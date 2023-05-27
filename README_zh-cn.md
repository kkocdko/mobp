# mobp

[![Release](https://img.shields.io/github/v/release/kkocdko/mobp?style=flat&color=293)](https://github.com/kkocdko/mobp/releases)
[![License](https://img.shields.io/github/license/kkocdko/mobp?style=flat&color=293)](LICENSE)

[English](README.md) • **简体中文**

微软办公套件后台进程。

> 仅支持Windows平台。
>
> 当前仅适配了三个常用组件（Word、Excel和PowerPoint）。

### 为什么需要这个程序？

在Mac OS上，Microsoft Office会保持后台运行。但在Windows上，单击关闭按钮就会导致进程退出。

如果你安装了较多插件，或是正在使用硬盘速度较慢的电脑，频繁地启动和关闭Microsoft Office将会浪费非常多时间。

此程序助你解决这一问题。

### 使用方法

1. 下载[已发布版本](https://github.com/kkocdko/mobp/releases)或者自行编译。
    * 如果下载失败，可以使用[备用链接](https://lanzoui.com/b0108u0pa)。

2. 确保Microsoft Office关联了`docx`，`xlk`和`pptx`这三个扩展名。

3. 运行`mobp.exe`。

### 注意事项

* 再次运行即可停止后台进程。

* 不建议与多标签页插件同时使用。

* 如果遇到 [窗口位置和大小异常](https://github.com/kkocdko/mobp/issues/1) 的情况，请注意：1. 尽量使用mobp自带的关闭功能。2. 先关闭mobp，手动打开Excel、Word 和 PowerPoint，将窗口最大化后关闭，然后再启动mobp。

### 兼容性

* Windows 10或更高版本。

* Microsoft Office 2013或更高版本。
