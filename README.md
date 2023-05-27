# mobp

[![Release](https://img.shields.io/github/v/release/kkocdko/mobp?style=flat&color=293)](https://github.com/kkocdko/mobp/releases)
[![License](https://img.shields.io/github/license/kkocdko/mobp?style=flat&color=293)](LICENSE)

**English** • [简体中文](README_zh-cn.md)

Microsoft Office Background Process.

> Windows only.
>
> Only three commonly used components (Word, Excel and PowerPoint) are currently supported.

### Why

On Mac OS, Microsoft Office keeps running in the background. But on Windows, clicking the close button will cause the process to exit.

If you have many plugins installed, or you are using a PC with a slow hard drive, starting and closing Office frequently will waste a lot of time.

This program will help you solve this problem.

### Usage

1. Download [release](https://github.com/kkocdko/mobp/releases) or compile yourself.
   - If the download fails, you can use [alternative link](https://lanzoui.com/b0108u0pa).

2. Make sure Microsoft Office is associated with the `docx`, `xlk` and `pptx` extension.

3. Run `mobp.exe`.

### Note

- Run again to stop background processes.

- Deprecated for use with multi-tab plugins.

- If the [window position and size were abnormal](https://github.com/kkocdko/mobp/issues/1) on your system, please: 1. Use mobp's built-in "close" function; 2. First close mobp, open Excel, Word and PowerPoint manually, maximize and then close the window, finally reopen mobp.

### Compatibility

- Windows 10 or higher.

- Microsoft Office 2013 or higher.
