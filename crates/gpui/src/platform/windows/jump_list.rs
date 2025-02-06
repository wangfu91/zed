use std::path::Path;

use anyhow::anyhow;
use windows::Win32::{
    System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER},
    UI::Shell::{
        Common::{IObjectArray, IObjectCollection},
        DestinationList, EnumerableObjectCollection, ICustomDestinationList, IShellLinkW,
        PropertiesSystem::{IPropertyStore, PSGetPropertyKeyFromName, PROPERTYKEY},
        ShellLink,
    },
};
use windows_core::{w, Interface, BSTR, HSTRING, PROPVARIANT};

pub(crate) fn add_to_jump_list(path: &Path) -> anyhow::Result<()> {
    log::info!("Adding {:?} to jump list", path);

    let jump_list: ICustomDestinationList =
        unsafe { CoCreateInstance(&DestinationList, None, CLSCTX_INPROC_SERVER)? };

    let items: IObjectCollection =
        unsafe { CoCreateInstance(&EnumerableObjectCollection, None, CLSCTX_INPROC_SERVER)? };

    let path_wstr: HSTRING = path.into();
    let exe_wstr: HSTRING = std::env::current_exe()?.as_os_str().into();
    log::info!("exe_wstr: {:?}", exe_wstr);
    let dir_wstr: HSTRING = Path::new(path)
        .file_name()
        .ok_or(anyhow!("path is not a directory"))?
        .into();
    let dir_wstr = BSTR::from_wide(dir_wstr.as_wide())?;

    // safety: FFI
    unsafe {
        let link = create_directory_link(exe_wstr, path_wstr, dir_wstr)?;
        items.AddObject(&link)?;
    }

    let array: IObjectArray = items.cast()?;
    unsafe { jump_list.AppendCategory(w!("Recent"), &array)? };
    unsafe { jump_list.CommitList()? };

    Ok(())
}

unsafe fn create_directory_link(
    exec_path: HSTRING,
    args: HSTRING,
    title: BSTR,
) -> anyhow::Result<IShellLinkW> {
    let link: IShellLinkW = CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER)?;
    link.SetPath(&exec_path)?;
    link.SetIconLocation(w!("explorer.exe"), 0)?; // // simulate folder icon
    link.SetArguments(&args)?; // folder path
    link.SetDescription(&args)?; // tooltip

    // the actual display string must be set as a property because IShellLink is primarily for shortcuts
    let title_value = PROPVARIANT::from(title);
    let mut title_key = PROPERTYKEY::default();
    PSGetPropertyKeyFromName(w!("System.Title"), &mut title_key)?;

    let store: IPropertyStore = link.cast()?;
    store.SetValue(&title_key, &title_value)?;
    store.Commit()?;

    Ok(link)
}
