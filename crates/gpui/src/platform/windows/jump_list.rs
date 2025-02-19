use std::path::{Path, PathBuf};

use anyhow::anyhow;
use windows::Win32::{
    Foundation::MAX_PATH,
    System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER},
    UI::Shell::{
        Common::{IObjectArray, IObjectCollection},
        DestinationList, EnumerableObjectCollection, ICustomDestinationList, IShellLinkW,
        PropertiesSystem::{IPropertyStore, PSGetPropertyKeyFromName, PROPERTYKEY},
        ShellLink,
    },
};
use windows_core::{w, Interface, BSTR, HSTRING, PROPVARIANT, PWSTR};

pub(crate) fn update_jump_list(paths: &[PathBuf]) -> anyhow::Result<()> {
    log::info!("========= Updating jump list with {:?}", paths);

    let mut recent = paths
        .iter()
        .map(|x| x.to_string_lossy().to_string())
        .collect::<Vec<_>>();

    // create a jump list
    let jump_list: ICustomDestinationList =
        unsafe { CoCreateInstance(&DestinationList, None, CLSCTX_INPROC_SERVER)? };

    // initialize the list and honor removals requested by the user
    let mut max_destinations = 0u32;
    let mut destination_path = vec![0u16; MAX_PATH as usize];
    unsafe {
        let removed_destinations: IObjectArray = jump_list.BeginList(&mut max_destinations)?;
        for i in 0..removed_destinations.GetCount()? {
            let removed_link: IShellLinkW = removed_destinations.GetAt(i)?;
            removed_link.GetArguments(&mut destination_path)?;
            let removed_path_wstr = PWSTR::from_raw(destination_path.as_mut_ptr());
            if !removed_path_wstr.is_null() {
                let removed_path = removed_path_wstr.to_string()?;
                recent.retain(|x| *x != removed_path);
            }
        }
    }

    let items: IObjectCollection =
        unsafe { CoCreateInstance(&EnumerableObjectCollection, None, CLSCTX_INPROC_SERVER)? };

    for path in recent {
        let path_wstr: HSTRING = HSTRING::from(&path);
        let exe_wstr: HSTRING = std::env::current_exe()?.as_os_str().into();
        let dir_wstr: HSTRING = Path::new(&path)
            .file_name()
            .ok_or(anyhow!("path is not a directory"))?
            .into();
        let dir_wstr = BSTR::from_wide(dir_wstr.as_wide())?;

        unsafe {
            let link = create_directory_link(exe_wstr, path_wstr, dir_wstr)?;
            items.AddObject(&link)?;
        }
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
