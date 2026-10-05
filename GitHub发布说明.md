# SlidePace GitHub 发布说明

GitHub 已发布版本为 `1.0.5`，源码放在 GitHub 仓库，用户安装文件放在 [Releases](https://github.com/Jimmy-1020/SlidePace/releases/tag/v1.0.5)。

1.0.5 新增“演示者侧显示计时器”独立开关，并保留 1.0.4 的 Esc 焦点修复。本次通过新标签 `v1.0.5` 发布，保留原 `v1.0.3`、`v1.0.4` 标签及发布记录。

1.0.4 的源码提交及标签为 `c2ee2fa62488b827c9563679076cf047f15ab074`，其已发布附件及内容保留。1.0.5 发布附件包括 `SlidePace-Setup.exe`、`SlidePace-Source.zip`、`SlidePace-1.0.5-Windows-x64.zip` 和 `SHA256SUMS.txt`；源码包和完整 Windows 包都包含对应的使用说明、需求文档与验证记录。

1.0.5 对应需求文档 v0.10：取消勾选“演示者侧显示计时器”后，保留演示者视图，仅隐藏该侧计时框，观众计时继续。设置会保存，放映中也可调整。8 组测试共 529 个断言通过；演示者内容自动避让仍待实现，实测范围见 [验证报告](./验证报告.md)。

项目仓库：[Jimmy-1020/SlidePace](https://github.com/Jimmy-1020/SlidePace)。

## 仓库内容

- `src/`：PowerPoint 插件和 EXE 安装器源码。
- `tests/`：核心、界面、安装隔离及 PowerPoint 集成验证。
- `tools/`、`build.cmd`、`SlidePace.sln`：构建和打包入口。
- `assets/`：产品图标。
- `docs/validation/`：文档引用的历史及当前版本截图、验证日志。
- `README.md`、需求文档、验证报告：中文项目说明与实际验收范围。

`.gitignore` 排除生成文件、安装包、Visual Studio 本地配置及临时测试数据。文档截图作为仓库文件保留。

## 上传到已有仓库

先克隆自己的 `SlidePace` 仓库，将源码 ZIP 解压后的文件放到仓库根目录，再提交并推送。也可通过 GitHub 网页的 **Add file → Upload files** 上传解压后的源码；上传 ZIP 文件本身不会自动展开成源码仓库。

如仓库已有许可证或其他文件，上传时保留这些文件。项目当前未自行指定开源许可证；可以在仓库中按自己的发布意愿选择。

## 从源码构建

构建端使用 Windows、Visual Studio 的 .NET 桌面开发工作负载、.NET Framework 4.8 目标包，以及项目所引用的 Office PIA。PowerPoint 实机测试需要已安装的 64 位 PowerPoint；GitHub 托管 runner 默认不包含 Office，本项目没有配置自动运行 Office 测试的工作流。

打开 `SlidePace.sln` 并生成 Release，或运行 `build.cmd`。也可在安装 Python 后运行：

```text
python tools/build.py
python tools/package.py
```

Python 构建脚本会通过 `vswhere` 或 PATH 查找 MSBuild，也支持通过 `SLIDEPACE_MSBUILD` 指定 MSBuild 的完整路径。发布所需截图已包含在 `docs/validation`，构建新安装包不依赖本机测试临时目录。

## 发布安装包

1. 在 GitHub 仓库打开 **Releases → Draft a new release**。
2. 核对准备发布的版本，为新版本创建对应标签及发布标题；例如 1.0.5 使用 `v1.0.5` 和 `SlidePace 1.0.5`。后续版本创建新标签，保留已有标签、发布记录及附件。
3. 上传 `dist` 中的 `SlidePace-Setup.exe`、`SlidePace-Source.zip`、`SHA256SUMS.txt`。如需完整说明，也上传中文 Markdown 文档及验证截图目录的压缩包。
4. 发布说明中写明适用于 Windows 64 位 PowerPoint，用户先保存并关闭 PowerPoint，再运行 EXE；列出本次改动及验证报告链接。

GitHub 会自动提供标签对应的源码 ZIP，项目生成的 `SlidePace-Source.zip` 则额外作为已核对的交付源码包。安装使用端无需 Visual Studio 或 Python。

## 1.0.3 已发布版本的文档修订记录

说明修订提交到 `main`，保留原 `v1.0.3` 标签。发布页注明本次文档提交，并更新 `SlidePace-Source.zip`、`SlidePace-1.0.3-Windows-x64.zip` 和 `SHA256SUMS.txt`，确保下载包中的说明与仓库一致。安装 EXE 的内容及校验和保持不变。

GitHub 自动生成的标签源码归档保留原发布时的文档；需要本次修订后的完整说明，请下载发布附件中的源码包或 Windows 完整包。替换后的下载文件须重新核对 SHA-256。
