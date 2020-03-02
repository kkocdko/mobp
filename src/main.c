#include <windows.h>

BOOL HideWindow(const char *window_title)
{
    HWND hwnd = FindWindow(NULL, window_title);
    if (hwnd)
    {
        SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSENDCHANGING);
        SetWindowLong(hwnd, GWL_EXSTYLE, WS_EX_NOACTIVATE);
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

int main()
{
    system(
        "cd.>\"%temp%\\_mobp.docx\""
        "&"
        "cd.>\"%temp%\\_mobp.csv\""
        "&"
        "cd.>\"%temp%\\_mobp.pptx\"");
    char temp_dir[512];
    GetTempPath(512, temp_dir);
    ShellExecute(NULL, NULL, "_mobp.docx", NULL, temp_dir, SW_SHOW);
    ShellExecute(NULL, NULL, "_mobp.csv", NULL, temp_dir, SW_SHOW);
    ShellExecute(NULL, NULL, "_mobp.pptx", NULL, temp_dir, SW_SHOW);
    BOOL word_hidden = FALSE;
    BOOL excel_hidden = FALSE;
    BOOL powerpoint_hidden = FALSE;
    while (!(word_hidden && excel_hidden && powerpoint_hidden))
    {
        // Think about the "Right to Left"?
        word_hidden = HideWindow("_mobp.docx - Word");
        excel_hidden = HideWindow("_mobp.csv - Excel");
        powerpoint_hidden = HideWindow("_mobp.pptx - PowerPoint");
        Sleep(5);
    }
    return 0;
}
