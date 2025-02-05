use windows::Win32::{
    System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER},
    UI::Shell::{
        Common::{IObjectArray, IObjectCollection},
        DestinationList, EnumerableObjectCollection, ICustomDestinationList, IShellLinkW,
    },
};

pub(crate) fn update_jump_list() -> anyhow::Result<()> {
    let jump_list: ICustomDestinationList =
        unsafe { CoCreateInstance(&DestinationList, None, CLSCTX_INPROC_SERVER)? };

    let mut min_slots = 0u32;
    let removed_list: IObjectArray = unsafe { jump_list.BeginList(&mut min_slots)? };

    for i in 0..min_slots {
        let mut removed_link: IShellLinkW = unsafe { removed_list.GetAt(i)? };
    }

    let items = IObjectCollection =
        unsafe { CoCreateInstance(&EnumerableObjectCollection, None, CLSCTX_INPROC_SERVER)? };

        

    Ok(())
}
