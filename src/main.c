#include <windows.h>
#include <stdio.h>

BOOL HideWindowByTitle(const char *window_title)
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

BOOL CloseWindowByTitle(const char *window_title)
{
    HWND hwnd = FindWindow(NULL, window_title);
    if (hwnd)
    {
        SendMessage(hwnd, WM_SYSCOMMAND, SC_CLOSE, 0);
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

int main(int argc, char *argv[])
{
    char *word_titles[] = {
        "_mobp - Microsoft Word",
        "_mobp.docx - Microsoft Word",
        "_mobp - Word",
        "_mobp.docx - Word",
        "Microsoft Word - _mobp",
        "Microsoft Word - _mobp.docx",
        "Word - _mobp",
        "Word - _mobp.docx"};
    char *excel_titles[] = {
        "_mobp - Microsoft Excel",
        "_mobp.csv - Microsoft Excel",
        "_mobp - Excel",
        "_mobp.csv - Excel",
        "Microsoft Excel - _mobp",
        "Microsoft Excel - _mobp.csv",
        "Excel - _mobp",
        "Excel - _mobp.csv"};
    char *powerpoint_titles[] = {
        "_mobp - Microsoft PowerPoint",
        "_mobp.pptx - Microsoft PowerPoint",
        "_mobp - PowerPoint",
        "_mobp.pptx - PowerPoint",
        "Microsoft PowerPoint - _mobp",
        "Microsoft PowerPoint - _mobp.pptx",
        "PowerPoint - _mobp",
        "PowerPoint - _mobp.pptx"};
    int word_titles_len = sizeof(word_titles) / sizeof(char *);
    int excel_titles_len = sizeof(excel_titles) / sizeof(char *);
    int powerpoint_titles_len = sizeof(powerpoint_titles) / sizeof(char *);
    BOOL word_closed = FALSE;
    BOOL excel_closed = FALSE;
    BOOL powerpoint_closed = FALSE;
    for (int i = 0; i < word_titles_len && !word_closed; i++)
    {
        word_closed = CloseWindowByTitle(word_titles[i]);
    }
    for (int i = 0; i < excel_titles_len && !excel_closed; i++)
    {
        excel_closed = CloseWindowByTitle(excel_titles[i]);
    }
    for (int i = 0; i < powerpoint_titles_len && !powerpoint_closed; i++)
    {
        powerpoint_closed = CloseWindowByTitle(powerpoint_titles[i]);
    }
    if (word_closed || excel_closed || powerpoint_closed)
    {
        printf("mobp is closed");
    }
    else
    {
        ShellExecute(
            NULL,
            NULL,
            "cmd.exe",
            "/c "
            "cd.>\"%temp%\\_mobp.docx\""
            "&"
            "cd.>\"%temp%\\_mobp.csv\""
            "&"
            "cd.>\"%temp%\\_mobp.pptx\""
            "&"
            "start \"\" \"%temp%\\_mobp.docx\""
            "&"
            "start \"\" \"%temp%\\_mobp.csv\""
            "&"
            "start \"\" \"%temp%\\_mobp.pptx\"",
            NULL,
            SW_HIDE);
        BOOL word_hidden = FALSE;
        BOOL excel_hidden = FALSE;
        BOOL powerpoint_hidden = FALSE;
        while (!(word_hidden && excel_hidden && powerpoint_hidden))
        {
            for (int i = 0; i < word_titles_len && !word_hidden; i++)
            {
                word_hidden = HideWindowByTitle(word_titles[i]);
            }
            for (int i = 0; i < excel_titles_len && !excel_hidden; i++)
            {
                excel_hidden = HideWindowByTitle(excel_titles[i]);
            }
            for (int i = 0; i < powerpoint_titles_len && !powerpoint_hidden; i++)
            {
                powerpoint_hidden = HideWindowByTitle(powerpoint_titles[i]);
            }
            Sleep(5);
        }
        printf("mobp is started");
    }
    if (argc == 1)
    {
        Sleep(2000);
    }
    return 0;
}
