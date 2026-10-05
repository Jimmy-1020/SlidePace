# SlidePace PowerPoint 计时插件

Windows PowerPoint 放映计时插件，当前版本 1.0.5。打开 [需求文档](./SlidePace-PowerPoint插件需求文档.md) 可查看完整功能定义。

1.0.5 新增“演示者侧显示计时器”独立开关，并保留 1.0.4 的 Esc 焦点修复。安装程序可从 [SlidePace 1.0.5 发布页](https://github.com/Jimmy-1020/SlidePace/releases/tag/v1.0.5) 下载，源码目录中的安装文件位于 `dist/SlidePace-Setup.exe`。

## 安装与使用

1. 保存并关闭所有 PowerPoint 窗口。
2. 双击安装程序 `SlidePace-Setup.exe`，等待自动安装完成；源码目录中的安装文件位于 `dist`。
3. 重新打开 PowerPoint，在“SlidePace 计时”选项卡选择一种模式。
4. 若选择倒计时，在“计时器设置”中设置时长，例如 `00:10:00`，并选择归零后“停止计时”或“继续顺计时”。
5. 启动放映，计时自动开始。每次打开 PowerPoint 都默认全部不选，因此需要先选择一种模式。

插件按当前 Windows 用户安装，使用端需要 64 位 Microsoft PowerPoint 桌面版和 .NET Framework 4.8。使用端无需 Visual Studio、Python 或 Office 开发工具。

## 计时与显示

- **顺计时**：显示经过时间，翻页持续累计。
- **倒计时**：归零后按设置停止在 `00:00:00`，或从零继续顺计时，例如 `00:00:03`，所有读数均不带负号。两种结束方式都改变数字颜色，默认红色，在“计时器设置”中可选其他颜色；默认继续顺计时。
- **系统时间**：显示当前本地时间，使用 24 小时制。
- **取消计时**：再次点击已选模式，暂停计时并移除所有屏幕的框体和附属入口。

框体使用半透明深色背景，平时只显示时间数字，宽度随文字调整，左右留白接近上下留白。显示范围包括实际幻灯片放映窗口和已打开的演示者视图，默认位于各自视图右上方；窗口放映时限制在窗口客户区，普通编辑界面不显示。

启用双屏演示者视图时，默认在观众放映画面和演示者视图各显示一份同步框体。若只需要观众侧计时，在“计时器设置”中取消勾选“演示者侧显示计时器”，点击“保存设置”：演示者视图保持打开，该侧计时框和收起入口隐藏，观众侧继续显示与计时。放映中也可从观众计时框的右键菜单打开设置调整；重新勾选后演示者侧显示当前同步读数，不重置计时。此开关默认开启，关闭或开启的选择会保留到下次使用。

未打开演示者视图时，只在实际放映窗口显示。单屏或复制模式共用的观众计时框不受该开关影响。两侧可分别设置位置，也可拖动数字区域调整位置；屏幕对应关系以实际放映窗口为准。全部模式未选时，两侧仍都不显示计时。

带有“下一个动画／下一张幻灯片”和备注区的界面是 PowerPoint 的演示者视图，属于放映过程。这里显示计时器符合“演讲者和观众都能看到”的需求；退出放映并返回普通编辑界面后，计时器隐藏。

演示者视图中的默认右上角以整个视图的客户区定位，可能覆盖下一动画或下一张幻灯片预览。可在“计时器设置 → 演讲者框体位置”选择其他角落，或拖动时间数字调整；观众侧的位置通过“观众框体位置”单独设置。也可关闭“演示者侧显示计时器”。当前版本仍没有自动避让预览／备注区域。

在“计时器设置”中选择数字字体和小／中／大字号，字体选择支持本机已安装字体并提供预览。字体、字号、倒计时时长、归零后行为、归零颜色和框体位置会保留；模式选择只在本次 PowerPoint 会话内生效。倒计时运行中需先暂停，再修改时长或结束行为。

顺计时和倒计时的三个按钮默认隐藏且不占框体高度，鼠标移入时临时展开按钮行，移出后恢复紧凑高度；靠近屏幕底部时向上展开，时间数字的位置保持稳定。运行时“开始”原位显示为“暂停”；暂停后使用“继续”，重置后使用“开始”。已完成的到零停止倒计时只能重置后开始新一轮。系统时间模式不显示这些按钮。鼠标悬停只影响所在屏幕，计时操作同步到双方。功能区只保留模式、“计时器设置”和使用说明，计时控制通过放映框体操作。

右键框体可打开设置、收起／展开或切换模式；再次点击菜单中已选模式也可取消。收起不停止计时。是否显示由模式选择决定：全部未选时不显示，选择模式后在放映中自动显示，无需额外开关。旧配置曾关闭显示也不会阻止计时器出现。

切回编辑界面或其他窗口覆盖某一放映视图时，该视图上的框体隐藏，计时继续；返回放映视图后恢复显示。结束放映自动暂停并隐藏所有框体。同一次 PowerPoint 会话内再次放映会继续保留的读数；已经完成且选择到零停止的倒计时保持零及结束颜色。需要从头计时，请在放映框体上点击“重置”，再点“开始”。重启 PowerPoint 后所有模式均未选中，重新选择计时模式后从初始读数开始。

选择倒计时后启动放映，无需先点击观众画面即可按 Esc 退出；点击演示者侧后按 Esc 也应结束整场放映，两侧计时框一起消失。1.0.4 使用无激活的原生置顶接口，保留 PowerPoint 的键盘焦点，不额外接管 Esc。

## 卸载与更新

在 Windows“设置 → 应用 → 已安装的应用”找到“SlidePace PowerPoint 计时插件”并卸载，也可运行安装目录内的 `Uninstall.exe`。卸载前先保存并关闭 PowerPoint；个人设置默认保留，可勾选删除。

更新时关闭 PowerPoint 并运行新的 EXE 安装包。安装失败会恢复原注册和原可用文件。卸载只清理插件标记的版本目录和自身设置文件，保留目录中的其他文件。

程序目录：`%LOCALAPPDATA%\SlidePace`。个人设置：`%APPDATA%\SlidePace\settings.json`。

## 开发与构建

使用 Visual Studio 打开 `SlidePace.sln`，选择 `Release` 并生成解决方案；也可运行根目录 `build.cmd`。开发端需要 .NET 桌面开发工具、.NET Framework 4.8 目标包和已安装的 Office PIA。Office 接口嵌入插件，发布时无需单独安装 PIA。

本机可使用 `python tools/build.py` 离线构建。产品图标已包含在 `assets`；需要重新绘制图标时运行 `python tools/generate_icon.py`（该辅助步骤需要 Pillow）。

生成 Release 后运行 `python tools/package.py`，在 `dist` 收集安装 EXE、中文文档、验证截图、源码 ZIP 和 SHA-256 校验文件。`build.cmd` 负责构建并输出安装 EXE，完整交付打包由 `package.py` 完成。

源码结构：

| 目录 | 用途 |
| --- | --- |
| `src/SlidePace.AddIn` | PowerPoint COM 入口、功能区、计时状态、双屏浮层和设置 |
| `src/SlidePace.Setup` | EXE 自动安装、用户级注册、回滚及卸载 |
| `tests` | 确定性计时测试、界面测试、安装隔离测试和 PowerPoint 集成验证 |
| `dist` | 使用端安装包与说明 |
| `artifacts` | 本机验证输出；不属于安装包 |

在工作目录运行以下测试：

```text
tests\bin\Release\SlidePace.Tests.exe --core artifacts\core
tests\bin\Release\SlidePace.Tests.exe --ui artifacts\ui
tests\bin\Release\SlidePace.Tests.exe --installer artifacts\installer
tests\bin\Release\SlidePace.Tests.exe --office artifacts\office
tests\bin\Release\SlidePace.Tests.exe --com-abi artifacts\com-abi
tests\bin\Release\SlidePace.Tests.exe --setup-ui artifacts\setup-ui
tests\bin\Release\SlidePace.Tests.exe --focus-ui artifacts\focus-ui
tests\bin\Release\SlidePace.Tests.exe --esc-native artifacts\esc-native
tests\bin\Release\SlidePace.Tests.exe --presenter-native artifacts\presenter-native
```

`--office` 使用测试创建的临时文稿，保留用户原有文稿；若已有活动放映，则中止测试。`--installer` 使用独立测试注册键，不注册到 Office。`--com-abi` 经非托管 COM 方法表调用生命周期接口，检查原生参数封送；`--setup-ui` 在工作目录与测试注册键中验证实际自动安装／卸载界面及环境预检。

`--focus-ui` 验证首次显示、刷新、悬停及重建框体时保留宿主子控件的焦点。`--esc-native` 使用独立临时 COM 标识，让 PowerPoint 在进程内加载本次构建的 Release DLL；现有正式插件临时取消计时，结束后恢复原模式并删除测试注册。测试创建临时文稿，以实际 F5 和 Esc 验证三种模式及未选模式的启动、演示者侧点击、双侧退出与浮层清理；启动后不主动激活放映窗口。已有活动放映时中止。`--esc-installed` 可对已安装版本执行倒计时及未选模式的同类诊断，1.0.3 的两条 Esc 路径已复现失败。测试均不会强制关闭用户 PowerPoint 或已有文稿。

`--presenter-native` 在上述原生 Esc 路径上增加演示者侧开关验证：放映中关闭／开启、预先关闭后启动，检查演示者视图 HWND 保留、观众侧数字像素可见、计时连续及 Esc 正常退出；四种模式状态分别验证。

`--host-load` 用于 Office 自主加载验证：先在 Visual Studio 中生成 Debug 配置的 AddIn 项目，再运行 `tests\bin\Release\SlidePace.Tests.exe --host-load artifacts\host-load`，使用临时工作目录和临时真实加载项注册检查 Ribbon、全屏放映和双屏浮层可见性。测试结束移除临时注册；Debug 配置仅在工作目录写入数据。正式发布使用 Release 配置。此验证要求 PowerPoint 中没有打开的文稿，已有正式 SlidePace 安装或活动放映时中止。1.0.3 本轮因检测到已有正式注册而未执行自主加载；已通过的真实放映测试使用新版 Release DLL，详情见验证报告。

当前 64 位 PowerPoint `16.0.20430.20118` 已完成首轮实机验证，其他版本不代表已验收。实际验证范围与尚未实测项目见 [验证报告](./验证报告.md)。

## GitHub 发布

源码仓库包含项目、构建脚本、需求文档和验证截图；安装 EXE 与源码 ZIP 通过 GitHub Releases 发布，`bin`、`obj`、`dist` 和测试临时输出不提交。具体上传及发布步骤见 [GitHub 发布说明](./GitHub发布说明.md)。
