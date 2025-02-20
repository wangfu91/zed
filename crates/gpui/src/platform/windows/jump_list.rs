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
use windows_core::{w, Interface, HSTRING, PROPVARIANT, PWSTR};

use crate::MenuItem;

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
            log::info!("====== removed_path={:?}", destination_path);
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
        let current_exe = std::env::current_exe()?;
        let dir_name = Path::new(&path)
            .file_name()
            .ok_or(anyhow!("path is not a directory"))?
            .to_str()
            .ok_or(anyhow!("path is not a valid unicode string"))?;

        unsafe {
            let link = create_shell_link(
                current_exe.as_path(),
                path.as_str(),
                path.as_str(),
                dir_name,
                &Path::new("explorer.exe"),
                0,
            )?;
            items.AddObject(&link)?;
        }
    }

    let array: IObjectArray = items.cast()?;
    unsafe { jump_list.AppendCategory(w!("Recent1"), &array)? };
    unsafe { jump_list.CommitList()? };

    Ok(())
}

pub(crate) fn add_tasks(menu_items: Vec<MenuItem>) -> anyhow::Result<()> {
    let jump_list: ICustomDestinationList =
        unsafe { CoCreateInstance(&DestinationList, None, CLSCTX_INPROC_SERVER)? };

    let items: IObjectCollection =
        unsafe { CoCreateInstance(&EnumerableObjectCollection, None, CLSCTX_INPROC_SERVER)? };

    let current_exe = std::env::current_exe()?;

    for menu_item in menu_items {
        if let MenuItem::Action { name, action, .. } = menu_item {
            log::info!(
                "Adding task to jump list: name={:?}, action={:?}",
                name,
                action
            );

            unsafe {
                let link = create_shell_link(
                    current_exe.as_path(),
                    "",
                    action.name(),
                    name.to_string().as_str(),
                    &current_exe.as_path(),
                    0,
                )?;
                items.AddObject(&link)?;
            }
        }
    }

    let array: IObjectArray = items.cast()?;
    let mut slots_visible: u32 = 0;
    let _removed: IObjectArray = unsafe { jump_list.BeginList(&mut slots_visible)? };
    log::info!("slots_visible={}", slots_visible);
    //unsafe { jump_list.AppendCategory(w!("Tasks"), &array)? };
    unsafe { jump_list.AddUserTasks(&array)? };
    unsafe { jump_list.CommitList()? };

    Ok(())
}

unsafe fn create_shell_link(
    program: &Path,
    args: &str,
    desc: &str,
    title: &str,
    icon_path: &Path,
    icon_index: i32,
) -> anyhow::Result<IShellLinkW> {
    let link: IShellLinkW = CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER)?;
    link.SetPath(&HSTRING::from(program))?;
    link.SetIconLocation(&HSTRING::from(icon_path), icon_index)?;
    link.SetArguments(&HSTRING::from(args))?; // path
    link.SetDescription(&HSTRING::from(desc))?; // tooltip

    // the actual display string must be set as a property because IShellLink is primarily for shortcuts
    let title_value = PROPVARIANT::from(title);
    let mut title_key = PROPERTYKEY::default();
    PSGetPropertyKeyFromName(w!("System.Title"), &mut title_key)?;

    let store: IPropertyStore = link.cast()?;
    store.SetValue(&title_key, &title_value)?;
    store.Commit()?;

    Ok(link)
}
