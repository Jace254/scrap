use self::ffi::*;
use std::{io, mem, ptr, slice};
use winapi::{
    HRESULT,
    IDXGIAdapter1,
    IDXGIFactory1,
    IDXGIOutput1,
    S_OK,
    UINT,
    DXGI_OUTPUT_DESC,
    LONG,
    DXGI_MODE_ROTATION,
    ID3D11Device,
    ID3D11DeviceContext,
    IDXGIOutputDuplication,
    D3D11_SDK_VERSION,
    D3D_DRIVER_TYPE_UNKNOWN,
    D3D_FEATURE_LEVEL_9_1,
    DXGI_ERROR_ACCESS_LOST,
    DXGI_ERROR_WAIT_TIMEOUT,
    DXGI_ERROR_INVALID_CALL,
    E_ACCESSDENIED,
    DXGI_ERROR_UNSUPPORTED,
    ID3D11Texture2D,
    DXGI_ERROR_NOT_CURRENTLY_AVAILABLE,
    DXGI_ERROR_SESSION_DISCONNECTED,
    TRUE,
    IDXGISurface,
    IDXGIResource,
    DXGI_RESOURCE_PRIORITY_MAXIMUM,
    D3D11_CPU_ACCESS_READ,
    D3D11_USAGE_STAGING,
    DXGI_OUTDUPL_POINTER_SHAPE_INFO,
    DXGI_OUTDUPL_FRAME_INFO,
    DXGI_OUTDUPL_POINTER_SHAPE_TYPE,
    DXGI_OUTDUPL_POINTER_SHAPE_TYPE_COLOR,
    DXGI_OUTDUPL_POINTER_SHAPE_TYPE_MASKED_COLOR,
    DXGI_OUTDUPL_POINTER_SHAPE_TYPE_MONOCHROME
};

mod ffi;

#[repr(C)]
struct CursorInfo {
    position: (i32, i32),
    shape: Vec<u8>,
    shape_info: DXGI_OUTDUPL_POINTER_SHAPE_INFO,
    visible: bool,
    who_updated_position_last: u32,
    last_time_stamp: i64,
}

pub struct Capturer {
    device: *mut ID3D11Device,
    context: *mut ID3D11DeviceContext,
    duplication: *mut IDXGIOutputDuplication,
    capture_mouse: bool,
    cursor_info: CursorInfo,
    fastlane: bool,
    surface: *mut IDXGISurface,
    data: *mut u8,
    len: usize,
    height: usize,
    width: usize,
    output_number: u32,
    offset_x: i32,
    offset_y: i32,
    desc: DXGI_OUTPUT_DESC,
}

impl Capturer {
    pub fn new(display: &Display, capture_mouse: bool) -> io::Result<Capturer> {
        let mut device = ptr::null_mut();
        let mut context = ptr::null_mut();
        let mut duplication = ptr::null_mut();
        let mut desc = unsafe { mem::uninitialized() };

        if unsafe {
            D3D11CreateDevice(
                &mut **display.adapter,
                D3D_DRIVER_TYPE_UNKNOWN,
                ptr::null_mut(),
                0,
                ptr::null_mut(),
                0,
                D3D11_SDK_VERSION,
                &mut device,
                &mut D3D_FEATURE_LEVEL_9_1,
                &mut context
            )
        } != S_OK {
            return Err(io::ErrorKind::Other.into());
        }

        let res = wrap_hresult(unsafe {
            (*display.inner).DuplicateOutput(
                &mut **device,
                &mut duplication
            )
        });

        if let Err(err) = res {
            unsafe {
                (*device).Release();
                (*context).Release();
            }
            return Err(err);
        }

        unsafe {
            (*duplication).GetDesc(&mut desc);
        }

        Ok(unsafe {
            let mut capturer = Capturer {
                device,
                context,
                duplication,
                fastlane: desc.DesktopImageInSystemMemory == TRUE,
                surface: ptr::null_mut(),
                height: display.height() as usize,
                width: display.width() as usize,
                data: ptr::null_mut(),
                len: 0,
                capture_mouse: capture_mouse,
                cursor_info: CursorInfo {
                    position: (0, 0),
                    shape: Vec::new(),
                    shape_info: mem::uninitialized(),
                    visible: false,
                    who_updated_position_last: 0,
                    last_time_stamp: 0,
                },
                output_number: 0, // Initialize this properly
                offset_x: 0, // Initialize this properly
                offset_y: 0, // Initialize this properly
                desc: display.desc.clone(),
            };
            let _ = capturer.load_frame(0);
            capturer
        })
    }

    unsafe fn load_frame(&mut self, timeout: UINT) -> io::Result<()> {
        let mut frame = ptr::null_mut();
        let mut info: DXGI_OUTDUPL_FRAME_INFO = mem::uninitialized();
        self.data = ptr::null_mut();

        wrap_hresult((*self.duplication).AcquireNextFrame(
            timeout,
            &mut info,
            &mut frame
        ))?;

        if self.capture_mouse {
            let mouse_update_time = info.LastMouseUpdateTime;
            if mouse_update_time != 0 {
                let update_position = if info.PointerPosition.Visible == 0 && self.cursor_info.who_updated_position_last != self.output_number {
                    false
                } else if info.PointerPosition.Visible != 0 && self.cursor_info.visible && 
                          self.cursor_info.who_updated_position_last != self.output_number && 
                          self.cursor_info.last_time_stamp > mouse_update_time {
                    false
                } else {
                    true
                };
    
                // update cursor position
                if update_position {
                    self.cursor_info.position = (
                        info.PointerPosition.Position.x + self.desc.DesktopCoordinates.left - self.offset_x,
                        info.PointerPosition.Position.y + self.desc.DesktopCoordinates.top - self.offset_y
                    );
                    self.cursor_info.who_updated_position_last = self.output_number;
                    self.cursor_info.last_time_stamp = mouse_update_time;
                    self.cursor_info.visible = info.PointerPosition.Visible != 0;
                }
    
                if info.PointerShapeBufferSize != 0 {
                    if info.PointerShapeBufferSize > self.cursor_info.shape.len() as u32 {
                        self.cursor_info.shape.resize(info.PointerShapeBufferSize as usize, 0);
                    }
                    let mut shape_size = 0;
                    wrap_hresult((*self.duplication).GetFramePointerShape(
                        info.PointerShapeBufferSize,
                        self.cursor_info.shape.as_mut_ptr() as *mut _,
                        &mut shape_size,
                        &mut self.cursor_info.shape_info,
                    ))?;
                }
            }
        }
        

        if self.fastlane {
            let mut rect = mem::uninitialized();
            let res = wrap_hresult(
                (*self.duplication).MapDesktopSurface(&mut rect)
            );

            (*frame).Release();

            if let Err(err) = res {
                Err(err)
            } else {
                self.data = rect.pBits;
                self.len = self.height * rect.Pitch as usize;
                Ok(())
            }
        } else {
            self.surface = ptr::null_mut();
            self.surface = self.ohgodwhat(frame)?;

            let mut rect = mem::uninitialized();
            wrap_hresult((*self.surface).Map(
                &mut rect,
                DXGI_MAP_READ
            ))?;

            self.data = rect.pBits;
            self.len = self.height * rect.Pitch as usize;
            Ok(())
        }
    }

    unsafe fn ohgodwhat(
        &mut self,
        frame: *mut IDXGIResource
    ) -> io::Result<*mut IDXGISurface> {
        let mut texture: *mut ID3D11Texture2D = ptr::null_mut();
        (*frame).QueryInterface(
            &IID_ID3D11TEXTURE2D,
            &mut texture as *mut *mut _ as *mut *mut _
        );

        let mut texture_desc = mem::uninitialized();
        (*texture).GetDesc(&mut texture_desc);

        texture_desc.Usage = D3D11_USAGE_STAGING;
        texture_desc.BindFlags = 0;
        texture_desc.CPUAccessFlags = D3D11_CPU_ACCESS_READ.0;
        texture_desc.MiscFlags = 0;

        let mut readable = ptr::null_mut();
        let res = wrap_hresult((*self.device).CreateTexture2D(
            &mut texture_desc,
            ptr::null(),
            &mut readable
        ));

        if let Err(err) = res {
            (*frame).Release();
            (*texture).Release();
            (*readable).Release();
            Err(err)
        } else {
            (*readable).SetEvictionPriority(DXGI_RESOURCE_PRIORITY_MAXIMUM);

            let mut surface = ptr::null_mut();
            (*readable).QueryInterface(
                &IID_IDXGISURFACE,
                &mut surface as *mut *mut _ as *mut *mut _
            );

            (*self.context).CopyResource(
                &mut **readable,
                &mut **texture
            );

            (*frame).Release();
            (*texture).Release();
            (*readable).Release();
            Ok(surface)
        }
    }

    pub fn frame<'a>(&'a mut self, timeout: UINT) -> io::Result<&'a [u8]> {
        unsafe {
            if self.fastlane {
                (*self.duplication).UnMapDesktopSurface();
            } else {
                if !self.surface.is_null() {
                    (*self.surface).Unmap();
                    (*self.surface).Release();
                    self.surface = ptr::null_mut();
                }
            }

            (*self.duplication).ReleaseFrame();

            self.load_frame(timeout)?;
            let frame = slice::from_raw_parts_mut(self.data, self.len);

            if self.capture_mouse && self.cursor_info.visible  {
                self.draw_cursor(frame);
            }
            Ok(slice::from_raw_parts(self.data, self.len))
        }
    }

    fn draw_cursor(&self, frame: &mut [u8]) {
        let (cursor_x, cursor_y) = self.cursor_info.position;
        let bytes_per_pixel = 4; // Assuming BGRA format
        let cursor_width = self.cursor_info.shape_info.Width as i32;
        let cursor_height = self.cursor_info.shape_info.Height as i32;
        let cursor_pitch = self.cursor_info.shape_info.Pitch as usize;
        let cursor_type = self.cursor_info.shape_info.Type;
        let frame_width = self.width as i32;
        let frame_height = self.height as i32;
        let shape_len = self.cursor_info.shape.len();
        
        let (hot_x, hot_y) = (
            self.cursor_info.shape_info.HotSpot.x as i32,
            self.cursor_info.shape_info.HotSpot.y as i32
        );
    
        for y in 0..cursor_height {
            for x in 0..cursor_width {
                let frame_x = cursor_x + x - hot_x;
                let frame_y = cursor_y + y - hot_y;
    
                if frame_x >= 0 && frame_y >= 0 && frame_x < frame_width && frame_y < frame_height {
                    let frame_index = (frame_y as usize * self.width + frame_x as usize) * bytes_per_pixel;
                    if frame_index + 3 < frame.len() {
                        let cursor_index = y as usize * cursor_pitch + x as usize * 4; // 4 bytes per pixel for color cursors
    
                        if cursor_index + 3 < shape_len {
                            match DXGI_OUTDUPL_POINTER_SHAPE_TYPE(cursor_type) {
                                DXGI_OUTDUPL_POINTER_SHAPE_TYPE_COLOR => {
                                    self.draw_color_cursor(frame, frame_index, cursor_index);
                                },
                                DXGI_OUTDUPL_POINTER_SHAPE_TYPE_MONOCHROME => {
                                    self.draw_monochrome_cursor(frame, frame_index, cursor_index, x);
                                },
                                DXGI_OUTDUPL_POINTER_SHAPE_TYPE_MASKED_COLOR => {
                                    self.draw_masked_color_cursor(frame, frame_index, cursor_index);
                                },
                                _ => {} // Unknown cursor type
                            }
                        }
                    }
                }
            }
        }
    }
    
    fn draw_color_cursor(&self, frame: &mut [u8], frame_index: usize, cursor_index: usize) {
        if cursor_index + 3 < self.cursor_info.shape.len() {
            let alpha = self.cursor_info.shape[cursor_index + 3] as u16;
            if alpha > 0 {
                for i in 0..3 {
                    if frame_index + i < frame.len() && cursor_index + i < self.cursor_info.shape.len() {
                        let cursor_color = self.cursor_info.shape[cursor_index + i] as u16;
                        let frame_color = frame[frame_index + i] as u16;
                        frame[frame_index + i] = ((alpha * cursor_color + (255 - alpha) * frame_color) / 255) as u8;
                    }
                }
                if frame_index + 3 < frame.len() {
                    frame[frame_index + 3] = 255; // Full opacity
                }
            }
        }
    }
    
    fn draw_monochrome_cursor(&self, frame: &mut [u8], frame_index: usize, cursor_index: usize, x: i32) {
        let byte_index = cursor_index / 8;
        let bit_index = 7 - (x % 8) as usize;
        if byte_index < self.cursor_info.shape.len() && byte_index + (self.cursor_info.shape_info.Height as usize / 2) < self.cursor_info.shape.len() {
            let and_mask = (self.cursor_info.shape[byte_index] >> bit_index) & 1;
            let xor_mask = (self.cursor_info.shape[byte_index + (self.cursor_info.shape_info.Height as usize / 2)] >> bit_index) & 1;
    
            if and_mask == 0 && xor_mask == 1 {
                // Invert the pixel
                for i in 0..3 {
                    if frame_index + i < frame.len() {
                        frame[frame_index + i] = 255 - frame[frame_index + i];
                    }
                }
            } else if and_mask == 0 && xor_mask == 0 {
                // Make the pixel black
                for i in 0..3 {
                    if frame_index + i < frame.len() {
                        frame[frame_index + i] = 0;
                    }
                }
            }
        }
    }
    
    fn draw_masked_color_cursor(&self, frame: &mut [u8], frame_index: usize, cursor_index: usize) {
        if cursor_index + 3 < self.cursor_info.shape.len() {
            let alpha = self.cursor_info.shape[cursor_index + 3] as u16;
            if alpha > 0 {
                for i in 0..3 {
                    if frame_index + i < frame.len() && cursor_index + i < self.cursor_info.shape.len() {
                        if self.cursor_info.shape[cursor_index + i] > 0 {
                            frame[frame_index + i] = self.cursor_info.shape[cursor_index + i];
                        }
                    }
                }
                if frame_index + 3 < frame.len() {
                    frame[frame_index + 3] = 255; // Full opacity
                }
            }
        }
    }    
}

impl Drop for Capturer {
    fn drop(&mut self) {
        unsafe {
            if !self.surface.is_null() {
                (*self.surface).Unmap();
                (*self.surface).Release();
            }
            (*self.duplication).Release();
            (*self.device).Release();
            (*self.context).Release();
        }
    }
}

pub struct Displays {
    factory: *mut IDXGIFactory1,
    adapter: *mut IDXGIAdapter1,
    /// Index of the CURRENT adapter.
    nadapter: UINT,
    /// Index of the NEXT display to fetch.
    ndisplay: UINT
}

impl Displays {
    pub fn new() -> io::Result<Displays> {
        let mut factory = ptr::null_mut();
        wrap_hresult(unsafe {
            CreateDXGIFactory1(&IID_IDXGIFACTORY1, &mut factory)
        })?;

        let mut adapter = ptr::null_mut();
        unsafe {
            // On error, our adapter is null, so it's fine.
            (*factory).EnumAdapters1(0, &mut adapter);
        };

        Ok(Displays {
            factory,
            adapter,
            nadapter: 0,
            ndisplay: 0
        })
    }

    // No Adapter => Some(None)
    // Non-Empty Adapter => Some(Some(OUTPUT))
    // End of Adapter => None
    fn read_and_invalidate(&mut self) -> Option<Option<Display>> {
        // If there is no adapter, there is nothing left for us to do.

        if self.adapter.is_null() {
            return Some(None);
        }

        // Otherwise, we get the next output of the current adapter.

        let output = unsafe {
            let mut output = ptr::null_mut();
            (*self.adapter).EnumOutputs(self.ndisplay, &mut output);
            output
        };

        // If the current adapter is done, we free it.
        // We return None so the caller gets the next adapter and tries again.

        if output.is_null() {
            unsafe {
                (*self.adapter).Release();
                self.adapter = ptr::null_mut();
            }
            return None;
        }

        // Advance to the next display.

        self.ndisplay += 1;

        // We get the display's details.

        let desc = unsafe {
            let mut desc = mem::uninitialized();
            (*output).GetDesc(&mut desc);
            desc
        };

        // We cast it up to the version needed for desktop duplication.

        let mut inner = ptr::null_mut();
        unsafe {
            (*output).QueryInterface(
                &IID_IDXGIOUTPUT1,
                &mut inner as *mut *mut _ as *mut *mut _
            );
            (*output).Release();
        }

        // If it's null, we have an error.
        // So we act like the adapter is done.

        if inner.is_null() {
            unsafe {
                (*self.adapter).Release();
                self.adapter = ptr::null_mut();
            }
            return None;
        }

        unsafe {
            (*self.adapter).AddRef();
        }

        Some(Some(Display { inner, adapter: self.adapter, desc }))
    }
}

impl Iterator for Displays {
    type Item = Display;
    fn next(&mut self) -> Option<Display> {
        if let Some(res) = self.read_and_invalidate() {
            res
        } else {
            // We need to replace the adapter.

            self.ndisplay = 0;
            self.nadapter += 1;

            self.adapter = unsafe {
                let mut adapter = ptr::null_mut();
                (*self.factory).EnumAdapters1(
                    self.nadapter,
                    &mut adapter
                );
                adapter
            };

            if let Some(res) = self.read_and_invalidate() {
                res
            } else {
                // All subsequent adapters will also be empty.
                None
            }
        }
    }
}

impl Drop for Displays {
    fn drop(&mut self) {
        unsafe {
            (*self.factory).Release();
            if !self.adapter.is_null() {
                (*self.adapter).Release();
            }
        }
    }
}

pub struct Display {
    inner: *mut IDXGIOutput1,
    adapter: *mut IDXGIAdapter1,
    desc: DXGI_OUTPUT_DESC
}

impl Display {
    pub fn width(&self) -> LONG {
        self.desc.DesktopCoordinates.right -
        self.desc.DesktopCoordinates.left
    }

    pub fn height(&self) -> LONG {
        self.desc.DesktopCoordinates.bottom -
        self.desc.DesktopCoordinates.top
    }

    pub fn rotation(&self) -> DXGI_MODE_ROTATION {
        self.desc.Rotation
    }

    pub fn name(&self) -> &[u16] {
        let s = &self.desc.DeviceName;
        let i = s.iter()
            .position(|&x| x == 0)
            .unwrap_or(s.len());
        &s[..i]
    }
}

impl Drop for Display {
    fn drop(&mut self) {
        unsafe {
            (*self.inner).Release();
            (*self.adapter).Release();
        }
    }
}

fn wrap_hresult(x: HRESULT) -> io::Result<()> {
    use std::io::ErrorKind::*;
    Err((match x {
        S_OK => return Ok(()),
        DXGI_ERROR_ACCESS_LOST => ConnectionReset,
        DXGI_ERROR_WAIT_TIMEOUT => TimedOut,
        DXGI_ERROR_INVALID_CALL => InvalidData,
        E_ACCESSDENIED => PermissionDenied,
        DXGI_ERROR_UNSUPPORTED => ConnectionRefused,
        DXGI_ERROR_NOT_CURRENTLY_AVAILABLE => Interrupted,
        DXGI_ERROR_SESSION_DISCONNECTED => ConnectionAborted,
        _ => Other
    }).into())
}