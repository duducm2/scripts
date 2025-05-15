#Include env.ahk ; Ensures IS_WORK_ENVIRONMENT is available

F12 & z::
{
    if (IS_WORK_ENVIRONMENT) {
        Run "C:\Users\fie7ca\Documents\Atalhos\WhatsApp.lnk"
    } else { ; Personal Environment
        Run "C:\Users\eduev\Documents\Atalhos\WhatsApp.lnk" ; Assuming .lnk for personal too, adjust if not
    }
}
