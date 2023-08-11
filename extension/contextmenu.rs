use windows::Win32::Foundation::{E_FAIL, E_UNEXPECTED};
use windows::Win32::System::Com::{IDataObject, FORMATETC, DVASPECT_CONTENT, TYMED_HGLOBAL};
use windows::Win32::System::Registry::HKEY;
use windows::Win32::UI::Shell::Common::ITEMIDLIST;
use windows::Win32::UI::Shell::{IContextMenu_Impl, CMINVOKECOMMANDINFO, CMF_DEFAULTONLY, DragQueryFileW, HDROP};
use windows::Win32::UI::WindowsAndMessaging::{HMENU, MB_ICONINFORMATION, MB_OK, InsertMenuItemW, MENUITEMINFOW, MIIM_STRING, MIIM_ID, MB_ICONERROR};
use windows::core::*;
use windows::Win32::UI::Shell::{IShellExtInit_Impl, IShellExtInit, IContextMenu};
use windows::Win32::Foundation::S_OK;
use windows::Win32::System::Ole::*;

use crate::msgBox;

#[implement(IShellExtInit, IContextMenu)]
#[derive(Default)]
pub struct MyRustExtension {

}

impl IShellExtInit_Impl for MyRustExtension {
    fn Initialize(&self, _pidlFolder: *const ITEMIDLIST, pdtObj: Option<&IDataObject>, _hkeyProgid: HKEY) -> Result<()> {
        // let folder = unsafe { pidlFolder.as_ref() };
        let Some(&obj) = pdtObj.as_ref() else { return E_UNEXPECTED.ok() };

        let format = FORMATETC {
            cfFormat: CF_HDROP.0,
            dwAspect: DVASPECT_CONTENT.0,
            lindex: -1,
            tymed: TYMED_HGLOBAL.0 as u32,
            ..Default::default()
        };
        unsafe {
            obj.QueryGetData(&format).ok()?;

            let medium = obj.GetData(&format)?;

            let hGlobal: HDROP = std::mem::transmute(medium.Anonymous.hGlobal);
            let file_index = 0xFFFFFFFF; // magic value
            let mut file_data = [0u16; 260];
            let count = DragQueryFileW(hGlobal, file_index, Some(&mut file_data));

            if count == 0 {
                return E_UNEXPECTED.ok()
            }
            let mut file_names = Vec::with_capacity(count as usize);

            for i in 0..count {
                DragQueryFileW(hGlobal, i, Some(&mut file_data));
                file_names.push(String::from_utf16_lossy(file_data.split(|x|*x==0).next().unwrap()))
            }

            msgBox(
                &wfmt!("My Rust Extension"),
                &wfmt!("count: {}\nfiles:{:?}", count, file_names),
                MB_OK|MB_ICONINFORMATION);
        }
        

        

        S_OK.ok()
    }
}

impl IContextMenu_Impl for MyRustExtension {
    fn QueryContextMenu(&self, hMenu: HMENU, indexMenu: u32, idCmdFirst: u32, _idCmdLast: u32, uFlags: u32) -> Result<()> {

        if (CMF_DEFAULTONLY & uFlags) != 0 {
            return S_OK.ok();
        }
        
        unsafe {

            let text: String = "i am a menu item 1".into();
            let mut bytes: Vec<_> = text.encode_utf16().chain(std::iter::once(0)).collect();
            let my_item = MENUITEMINFOW {
                cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                fMask: MIIM_STRING | MIIM_ID,
                wID: idCmdFirst, // idCmdFirst <= wID <=  idCmdLast
                dwTypeData: PWSTR(bytes.as_mut_ptr()),
                ..Default::default()
            };
            let success = InsertMenuItemW(
                hMenu,
                indexMenu,
                true,
                (&my_item) as * const _
            );

            if !success.as_bool() {
                msgBox(
                    &wfmt!("My Rust Extension"),
                    &wfmt!("InsertMenuItemW err"),
                    MB_OK|MB_ICONERROR);
                return Err(Error::from_win32())
            }

            let max_wID = my_item.wID as i32;
            Err(HRESULT(max_wID - idCmdFirst as i32 + 1).into())
        }
    }

    fn InvokeCommand(&self, lpcmi: *const CMINVOKECOMMANDINFO) -> Result<()> {
        let Some(&lpcmi) = (unsafe { lpcmi.as_ref() }) else { return E_UNEXPECTED.ok() };

        if (lpcmi.lpVerb.as_ptr() as usize) & 0xFF != 0 {
            return E_FAIL.ok();
        }

        S_OK.ok()
    }

    fn GetCommandString(&self, _idCmd: usize, _uFlags: u32, _: *const u32, _pszName: PSTR, _cchName: u32) -> Result<()> {
        S_OK.ok()
    }
}
