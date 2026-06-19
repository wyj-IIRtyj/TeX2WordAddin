# TeX2Word — LaTeX 公式转 Word 原生公式插件

<p align="center">
  <img src="assets/icon-128.png" width="96" alt="TeX2Word Icon" />
</p>

<p align="center">
  <strong>将 LaTeX 数学公式一键转换为 Microsoft Word 原生 OMML 公式</strong><br/>
  支持整页扫描批量替换 &amp; 智能粘贴混合文本
</p>

<p align="center">
  <a href="https://wyj-iirtyj.github.io/TeX2WordAddin/">在线部署地址</a> ·
  <a href="#安装指南">安装指南</a> ·
  <a href="#使用说明">使用说明</a>
</p>

---

## 功能特性

- **整页扫描模式 (Scheme A)**: 自动检索文档中所有 LaTeX 公式（`$$...$$`、`$...$`、`\(...\)`、`\[...\]`），列出预览后支持逐个或批量原位转换
- **智能粘贴模式 (Scheme B)**: 粘贴包含 LaTeX 的混合文本，自动分离公式并渲染为 Word 原生公式插入光标位置
- **完整公式支持**: 基于 MathJax 引擎，支持绝大多数 LaTeX 数学语法
- **日志调试面板**: 内置错误日志收集与一键复制，便于排查问题

## 技术架构

```
LaTeX → MathJax (tex→MathML) → XSLT (MathML→OMML) → OOXML Package → Word 原生公式
```

---

## 安装指南

> **前提条件**: 需要 Microsoft 365 订阅版 Word（桌面端或 Web 端），或 Office 2021 及以上版本。永久授权的 Office 2019 及更早版本不支持侧载加载项。

### 步骤 1: 下载 manifest 配置文件

下载本仓库 [`install/manifest.xml`](install/manifest.xml) 文件到本地。

> 你可以直接点击上方链接进入 GitHub 文件页面，然后点击 **Raw** 按钮后右键"另存为"下载，或者使用以下命令：
>
> ```bash
> curl -L -o manifest.xml https://raw.githubusercontent.com/wyj-IIRtyj/TeX2WordAddin/main/install/manifest.xml
> ```

---

### macOS Word 安装

1. **打开 manifest 目录**

   在 Finder 中按 `⌘ + Shift + G`（前往文件夹），输入以下路径：

   ```
   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   ```

   > 如果 `wef` 文件夹不存在，请手动创建。

2. **放入 manifest 文件**

   将下载的 `manifest.xml` 复制到 `wef` 文件夹中。

3. **重启 Word**

   完全退出 Word（`⌘ + Q`），然后重新打开。

4. **启用插件**

   打开任意 Word 文档后，在菜单栏点击：

   **插入 → 加载项 → 我的加载项**（或 **Insert → Add-ins → My Add-ins**）

   在弹出窗口中选择 **共享文件夹**（Shared Folder），找到 **TeX2WordAddin** 并点击添加。

5. **使用插件**

   添加成功后，在 **开始**（Home）选项卡右侧会出现 **TeX2Word** 按钮，点击即可打开插件面板。

---

### Windows Word 安装

#### 方法 A: 共享文件夹侧载（推荐）

1. **创建共享文件夹**

   在本地创建一个文件夹，例如：

   ```
   C:\OfficeAddins\
   ```

2. **放入 manifest 文件**

   将下载的 `manifest.xml` 复制到此文件夹。

3. **设置为共享文件夹**

   右键该文件夹 → **属性** → **共享** 选项卡 → 点击 **共享** → 确认共享（可选"读取"权限即可）。记下共享路径，如 `\\你的电脑名\OfficeAddins`。

4. **在 Word 中信任此目录**

   打开 Word → **文件** → **选项** → **信任中心** → **信任中心设置** → **受信任的加载项目录**

   在 **目录 URL** 中输入共享路径（如 `\\你的电脑名\OfficeAddins`），点击 **添加目录**，勾选 **显示在菜单中**，点击确定。

5. **重启 Word 并添加插件**

   重启 Word，然后进入 **插入** → **获取加载项** → **共享文件夹**，找到 TeX2WordAddin 并添加。

#### 方法 B: 网络共享目录

适用于团队部署，将 `manifest.xml` 放入组织的网络共享目录，由 IT 管理员通过集中部署分发。

---

### iOS / iPad Word 安装

> ⚠️ iOS/iPad 上的 Word 加载项功能较为有限，需要 Microsoft 365 订阅。

1. **打开 Word 应用**，打开或新建一个文档。

2. **点击右上角的 "..." 菜单**（更多选项）。

3. **选择"加载项"**（Add-ins）。

4. **选择"我的加载项"** → **"管理我的加载项"**。

5. 遗憾的是，iOS/iPad 版 Word **不支持直接侧载自定义 manifest 文件**。你需要通过以下方式之一使用：

   **方式一：通过 Microsoft 365 管理中心集中部署**
   
   如果你有 Microsoft 365 管理员权限：
   
   - 登录 [Microsoft 365 管理中心](https://admin.microsoft.com)
   - 前往 **设置** → **集成应用** → **加载项**
   - 点击 **上传自定义应用** → 选择 `manifest.xml`
   - 部署完成后，该加载项将自动出现在所有设备的 Word 中（包括 iPad）

   **方式二：使用 Word Online**
   
   - 在 iPad 浏览器中打开 [Word Online](https://www.office.com)
   - 使用 **插入** → **加载项** → **上传我的加载项** 功能上传 `manifest.xml`

---

### Word Online (Web 版) 安装

1. 在浏览器中打开 [Word Online](https://www.office.com)，打开或新建文档。
2. 点击 **插入** → **加载项** → **管理我的加载项** → **上传我的加载项**。
3. 点击 **浏览**，选择下载的 `manifest.xml` 文件。
4. 上传成功后，在 **开始** 选项卡中将出现 **TeX2Word** 按钮。

---

## 使用说明

### 整页扫描模式 (Scheme A)

1. 在 Word 文档中正常书写或粘贴包含 LaTeX 公式的文本，例如：

   ```
   由公式 $$E = mc^2$$ 可知，能量与质量...
   行内公式 $\alpha + \beta = \gamma$ 也支持。
   ```

2. 点击 **TeX2Word** 按钮打开插件面板，默认在 **整页扫描** 选项卡。

3. 点击 **"检索文档公式"** 按钮，插件将自动扫描文档中所有 LaTeX 公式。

4. 扫描完成后，公式列表将显示在面板中，每个条目包含：
   - 序号 + 类型标签（块级/行内）
   - MathJax 渲染的实时预览
   - 独立的"转换"按钮

5. 在顶部操作栏中选择：
   - **取消**: 清除扫描列表
   - **转换所选 (1个)**: 点选某个公式后，单独转换
   - **转换全部**: 批量一键转换所有公式

### 智能粘贴模式 (Scheme B)

1. 切换到 **智能粘贴** 选项卡。
2. 在文本框中粘贴包含 LaTeX 公式的混合文本。
3. 点击 **"解析并混合注入当前光标"**，文本和公式将自动注入到 Word 文档光标位置。

### 日志调试

如遇到转换失败，可切换到 **日志记录** 选项卡：
- 仅显示警告和错误级别的日志
- 点击 **"复制关键错误日志"** 可一键复制到剪贴板，方便反馈问题

---

## 本地开发

```bash
# 克隆仓库
git clone https://github.com/wyj-IIRtyj/TeX2WordAddin.git
cd TeX2WordAddin

# 安装依赖
npm install

# 启动开发服务器 (https://localhost:3000)
npm run dev-server

# 在 Word 中侧载开发版本
npm run start

# 生产构建
npm run build
```

---

## 许可证

MIT License
