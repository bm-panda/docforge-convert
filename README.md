# 文档格式转换

> 一键把选中的文档、表格、演示、电子书转换为你需要的格式。本脚本基于开源项目 **DocForge** 二次开发（见文末致谢），接入不忙脚本盒子并本地离线运行，文件不出本机。

---

## 核心特性

- **文档互转**：Word(docx) / ODT / RTF / Markdown / HTML / 纯文本 与 PDF 互转
- **表格互转**：xlsx / xls / csv / tsv / ods ↔ xlsx / csv / tsv / json / md / html / pdf / docx / ods
- **演示转换**：pptx / ppt / odp → PDF / TXT / Markdown / HTML
- **PDF 全能**：PDF → Word / 图片(PNG/JPG) / 文本 / Markdown / HTML
- **电子书**：EPUB → TXT / Markdown / HTML / PDF / DOCX
- **高保真 Office→PDF**：本机安装了 Microsoft Office 时自动走 Word/Excel/PowerPoint 引擎导出，未安装则回退普通转换
- **批量处理**：一次选择多个文件统一转换
- **本地离线**：转换在本机完成，不上传云端

---

## 脚本演示

![文档格式转换演示](./imgs/demo.gif "文档格式转换演示")



---

## 安装方法

本脚本的运行环境、依赖组件及图形化窗口均依赖“不忙脚本盒子”，建议仅通过盒子安装和使用本脚本。

> 前置条件：安装好 [不忙脚本盒子](https://www.bm-box.cn)

* 打开盒子 → **脚本市场** → 搜索 **「文档格式转换」**
* 点击「安装」（首次会自动安装 Python 依赖，需联网）

---

## 使用方式

启动后会弹出盒子渲染的参数窗口：选择源文件 → 选择目标格式 → 点「执行」。

* 直接启动
  * 在盒子里点击脚本卡片，在弹出窗口中点「源文件」右侧的选择按钮挑选文件（可多选），选择「目标格式」，点「执行」
* 右键启动
  * 在文件管理器中选中一个或多个文件，右键 → 选择「文档格式转换」
* 快捷键启动
  * 为脚本设置全局快捷键，选中文件后按下即可

---

## 💡 使用提示

- **目标格式要与源文件匹配**：
例如 docx 可转 pdf/md/html/txt；PDF 可转 docx/png/txt；表格可转 csv/xlsx/pdf 等。不支持组合会明确提示失败原因。
- **输出目录**
留空时，结果默认保存在源文件所在目录（同名会加序号）。
- **PDF → 图片** 
会输出到 `<文件名>_to_png`（或 `_to_jpg`）文件夹，多页分别为 `_p1/_p2…`。
- **Office → PDF 高保真** 
需要本机安装 Microsoft Office；未安装时 Word/PPT/Excel → PDF 会退化为“纯文字排版”。
- **旧版 .doc / .ppt** 
必须有 Microsoft Office 才能转换。
- **本脚本无图形界面**：
逐文件的转换进度、成功/失败原因与总耗时都显示在命令行窗口中；转换结束后窗口**暂停 5 秒**再自动关闭，方便查看结果。

---

## 节点脚本(开发者阅读)

本脚本声明了 `node`，可被其他脚本通过 `/api/link` 联动调用。

```json
// data.input_files 由盒子填入选中文件；params 由调用方传入
{ "target": "pdf", "output_dir": "D:/out" }
```

返回信封（写入 `environment.output_json`）：

```json
{ "code": 0, "msg": "ok", "results": "[{\"source\":\"...\",\"target\":\"...\",\"status\":\"success\",\"message\":\"...\",\"outputs\":[\"...\"]}]" }
```

- 业务键 `results`：结果 JSON 字符串；`code = 0` 全部成功，`1` 存在失败，`msg` 为摘要或错误。

---

## 🖥️ 环境要求

| 项目 | 要求 |
| --- | --- |
| **操作系统** | Windows 10/11（64 位） |
| **运行环境** | 不忙脚本盒子（内置 Python 3.10+） |
| **可选组件** | Microsoft Office（用于 Office→PDF 高保真与 doc/ppt 转换） |

---

## 📜 开源致谢

本脚本基于开源项目 **DocForge** 二次开发（作者 **binghai-fly**，MIT 许可，仓库 <https://github.com/binghai-fly/DocForge>），特此致谢；原项目 `LICENSE` 已随附。
