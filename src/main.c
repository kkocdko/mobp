#include <stdio.h>
#include <windows.h>

HWND QueryWindow(char *class_name, char *title_keyword)
{
    char current_title[64];
    HWND hwnd = NULL;
    HWND hwnd_child_after = NULL;
    while ((hwnd = FindWindowEx(NULL, hwnd_child_after, class_name, NULL)))
    {
        GetWindowText(hwnd, current_title, 64);
        if (strstr(current_title, title_keyword) != 0)
        {
            break;
        }
        hwnd_child_after = hwnd;
    }
    return hwnd;
}

BOOL CloseWindowByQuery(char *class_name, char *title_keyword)
{
    HWND hwnd = QueryWindow(class_name, title_keyword);
    if (hwnd)
    {
        PostMessage(hwnd, WM_SYSCOMMAND, SC_CLOSE, 0);
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

BOOL HideWindowByQuery(char *class_name, char *title_keyword)
{
    HWND hwnd = QueryWindow(class_name, title_keyword);
    if (hwnd)
    {
        SetWindowPos(hwnd, HWND_BOTTOM, -512, -512, 0, 0, SWP_NOSENDCHANGING);
        SetWindowLong(hwnd, GWL_EXSTYLE, WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW);
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

int main(int argc, char *argv[])
{
    char *title_keyword = "_mobp";
    char *word_class_name = "OpusApp";
    char *excel_class_name = "XLMAIN";
    char *powerpoint_class_name = "PPTFrameClass";
    BOOL word_closed = CloseWindowByQuery(word_class_name, title_keyword);
    BOOL excel_closed = CloseWindowByQuery(excel_class_name, title_keyword);
    BOOL powerpoint_closed = CloseWindowByQuery(powerpoint_class_name, title_keyword);
    if (word_closed || excel_closed || powerpoint_closed)
    {
        puts("mobp is closed");
    }
    else
    {
        ShellExecute(
            NULL,
            NULL,
            "cmd",
            "/c "
            "cd /d %temp%"
            "&"
            "cd.>_mobp.docx"
            "&"
            "cd.>_mobp.xlk"
            "&"
            "cd.>_mobp.pptx"
            "&"
            "start /b cmd /c start _mobp.docx"
            "&"
            "start /b cmd /c start _mobp.xlk"
            "&"
            "start /b cmd /c start _mobp.pptx",
            NULL,
            SW_HIDE);
        BOOL word_hidden = FALSE;
        BOOL excel_hidden = FALSE;
        BOOL powerpoint_hidden = FALSE;
        while (!(word_hidden && excel_hidden && powerpoint_hidden))
        {
            if (!word_hidden)
            {
                word_hidden = HideWindowByQuery(word_class_name, title_keyword);
            }
            if (!excel_hidden)
            {
                excel_hidden = HideWindowByQuery(excel_class_name, title_keyword);
            }
            if (!powerpoint_hidden)
            {
                powerpoint_hidden = HideWindowByQuery(powerpoint_class_name, title_keyword);
            }
            Sleep(5);
        }
        puts("mobp is started");
    }
    if (argc == 1)
    {
        Sleep(2000);
    }
    return 0;
}
