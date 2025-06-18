# canvas-editor-main 项目结构与文件说明

本项目为基于Canvas/SVG的富文本编辑器，以下为各主要目录和文件的结构、功能及其依赖说明：

---

## 1. 根目录文件

- **package.json**
  - 位置：根目录
  - 功能：项目依赖、脚本、元信息配置
  - 主要引用：prismjs、typescript、vite、eslint、cypress、vitepress、vue等

- **README.md**
  - 位置：根目录
  - 功能：项目说明文档

- **vite.config.ts**
  - 位置：根目录
  - 功能：Vite 构建工具配置
  - 主要引用：vite、@vitejs/plugin-vue等

- **tsconfig.json**
  - 位置：根目录
  - 功能：TypeScript 配置

- **index.html**
  - 位置：根目录
  - 功能：项目入口HTML

- **yarn.lock / package-lock.json**
  - 位置：根目录
  - 功能：依赖锁定文件

---

## 2. src 目录

### 2.1 src/main.ts
- 位置：src/main.ts
- 功能：项目主入口，初始化编辑器，绑定UI事件
- 主要引用：./editor、./mock、./components/dialog/Dialog、./components/signature/Signature、prismjs等

### 2.2 src/mock.ts
- 位置：src/mock.ts
- 功能：模拟数据和配置

### 2.3 src/style.css
- 位置：src/style.css
- 功能：全局样式

### 2.4 src/vite-env.d.ts
- 位置：src/vite-env.d.ts
- 功能：Vite环境类型声明

### 2.5 src/assets/
- 位置：src/assets/
- 功能：静态资源（如图片、快照等）

### 2.6 src/utils/
- 位置：src/utils/
- 功能：通用工具函数
- 主要文件：
  - index.ts：工具函数集合
  - prism.ts：prismjs相关处理

### 2.7 src/components/
- 位置：src/components/
- 功能：UI组件
- 主要子目录：
  - dialog/：对话框组件（Dialog.ts, dialog.css）
  - signature/：签名组件（Signature.ts, signature.css）

### 2.8 src/plugins/
- 位置：src/plugins/
- 功能：插件扩展
- 主要子目录：markdown/、copy/

---

## 3. src/editor 编辑器核心

### 3.1 src/editor/index.ts
- 位置：src/editor/index.ts
- 功能：编辑器主类，导出Editor及相关类型、常量、枚举
- 主要引用：core、interface、dataset、utils等

### 3.2 src/editor/core/
- 位置：src/editor/core/
- 功能：核心功能模块
- 主要子模块：
  - draw/：绘图与渲染（Draw.ts，frame/、richtext/、particle/等）
  - command/：命令模式（Command.ts, CommandAdapt.ts）
  - listener/：事件监听（Listener.ts）
  - register/：注册机制（Register.ts）
  - shortcut/：快捷键（Shortcut.ts, keys/）
  - plugin/：插件机制（Plugin.ts）

### 3.3 src/editor/interface/
- 位置：src/editor/interface/
- 功能：类型和接口定义
- 主要文件：Editor.ts、Element.ts、Draw.ts、Control.ts等

### 3.4 src/editor/dataset/
- 位置：src/editor/dataset/
- 功能：常量和枚举
- 主要子目录：constant/、enum/

### 3.5 src/editor/utils/
- 位置：src/editor/utils/
- 功能：编辑器专用工具函数
- 主要文件：element.ts（元素处理）、option.ts（配置合并）、clipboard.ts（剪贴板）等

### 3.6 src/editor/types/
- 位置：src/editor/types/
- 功能：类型声明

---

## 4. 主要文件引用关系举例

- **src/main.ts**
  - 引用：./editor（主编辑器）、./mock（模拟数据）、./components/dialog/Dialog、prismjs
- **src/editor/index.ts**
  - 引用：core/draw/Draw、core/command/Command、core/listener/Listener、core/register/Register、core/plugin/Plugin、utils/element等
- **src/editor/core/draw/Draw.ts**
  - 引用：core/cursor/Cursor、core/event/CanvasEvent、core/history/HistoryManager、core/observer/ScrollObserver、dataset/enum/Editor、interface/Draw等
- **src/editor/core/command/Command.ts**
  - 引用：core/command/CommandAdapt、core/draw/Draw、interface/Command等
- **src/components/dialog/Dialog.ts**
  - 引用：dialog.css

---

## 5. 依赖的主要外部库

- prismjs：代码高亮
- typescript：类型系统
- vite：构建工具
- eslint：代码质量检查
- cypress：端到端测试
- vitepress：文档生成
- vue：部分UI或文档

---

> 以上为canvas-editor-main项目的主要文件、目录、功能及引用关系说明。详细代码请查阅各目录下具体文件。