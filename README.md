<p>MinStart – Windows 最小化启动快捷方式工具<br>
MinStart 是一个用于 Windows 的轻量级启动辅助工具，用来在开机时自动启动指定快捷方式（.lnk），并将这些程序的窗口最小化到后台，避免桌面被一堆启动窗口覆盖。<br>
<br>

应用场景：<br>
1.软件自带的开机启动偶发性失灵<br>
2.使用系统启动文件夹启动快捷方式时，开机启动的应用会占满屏幕需要逐个最小化<br>

功能特点<br>
1.扫描指定路径所有 .lnk 文件<br>
2.使用 WScript.Shell 解析每个快捷方式对应的目标程序路径，提取 exe 文件名作为目标进程名集合并发启动所有 .lnk<br>
3.枚举系统进程，过滤出目标 exe 名<br>
4.对这些进程查找可见顶级窗口并调用 ShowWindow(SW_MINIMIZE) 进行最小化<br>
5.监控时间结束后脚本自动退出<br>
<br>
支持通过 JSON 配置文件自定义：<br>
配置文件路径 C:\Users\Public\minstart.json<br>
1.快捷方式目录 lnk_dir<br>
2.监控总时长 monitor_seconds<br>
3.扫描间隔 scan_interval<br>
<br>
编译为单文件 exe（可选）：<br>
如果希望在没有 Python 环境的机器上使用，可以通过 PyInstaller 打包：<br>
<br>
<pre>
pyinstaller -F -w -i minstart.ico --runtime-tmpdir "C:\Users\Public\tmp" minstart.py</pre></p>
