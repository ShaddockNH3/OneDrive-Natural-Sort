# OneDrive Natural Sort

让 OneDrive / SharePoint 原生文件列表按“自然排序”显示文件名里的数字，避免默认字符串排序把 `1, 10, 2` 排成奇怪顺序。

## 功能

- 支持升序自然排序：`1, 2, 10`
- 支持降序自然排序：`10, 2, 1`
- 支持 Chrome 和 Microsoft Edge 的 Tampermonkey / Violentmonkey
- 匹配个人 OneDrive 与常见 SharePoint / OneDrive for Business 页面
- 拦截 OneDrive 原生 `RenderListDataAsStream` 数据响应
- 在 OneDrive 原生虚拟列表渲染前改写数据顺序，不创建额外文件列表页面

## 从 GitHub 安装

1. 在 Chrome 或 Edge 安装 Tampermonkey。
2. 如果使用 Edge，打开扩展详情页并启用 `允许用户脚本`。Tampermonkey 面板顶部不应再提示这一步。
3. 打开脚本安装链接：

   <https://raw.githubusercontent.com/ShaddockNH3/OneDrive-Natural-Sort/main/onedrive-natural-sort.user.js>

4. Tampermonkey 会自动打开安装页面，点击安装。
5. 刷新 OneDrive 页面。

## 使用

脚本启用后，页面右下角会出现三个按钮：

- `A->Z`：原生列表按数字自然升序
- `Z->A`：原生列表按数字自然降序
- `Reload`：刷新 OneDrive，让新方向或拦截器重新应用到原生列表数据

脚本必须在 OneDrive 请求列表数据之前运行，所以使用了 `@run-at document-start`。如果刚安装后页面没有变化，点 `Reload` 或刷新页面。

## 原理

OneDrive 新版列表是虚拟列表，只改 DOM 会变成局部排序。本脚本改为拦截 SharePoint/OneDrive 的 `RenderListDataAsStream` 响应。

脚本会把原始列表请求改成一次取更多行，然后在响应被 OneDrive 前端读取之前，按 `FileLeafRef` / 文件名进行自然排序。这样 OneDrive 仍然使用自己的原生列表和虚拟滚动，只是拿到的数据顺序已经是自然排序后的顺序。

## 注意

这是对 OneDrive 内部接口的用户脚本级改写。微软如果调整 `RenderListDataAsStream` 的响应结构、请求方式或虚拟列表逻辑，脚本可能需要更新。

## 开发

语法检查：

```powershell
node --check .\onedrive-natural-sort.user.js
```
