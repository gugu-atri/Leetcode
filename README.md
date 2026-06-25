# LeetCode Hot 100 晨读背诵笔记

> 力扣 Hot 100 题解 · 思路速记 · 背诵口诀 · Java

官方分类顺序整理，每题包含：题目描述 + 示例 + ✎ 学霸笔记（口诀化速记）。
适合每天早晨朗读记忆，把题解思路『吟唱』成肌肉记忆。

---

## 📚 三种产物

| 产物 | 路径 | 角色 | 编辑方式 |
|------|------|------|----------|
| **Word 文档** | [`文档版/`](./文档版/) | ⭐ **唯一信息源** | Word / WPS 手动编辑 |
| **PDF** | [`PDF版/`](./PDF版/) | docx 渲染产物 | `bash scripts/build.sh` |
| **Markdown** | [`solution/`](./solution/) | docx 拆分产物（搜索 / git diff / 阅读） | `python3 scripts/split-by-problem.py` |

> **核心原则**：docx 是『真理之源』。所有内容先在 docx 里编辑，pdf / md 只是从 docx 派生。
> 这样能保留你在 Word 里精心调好的排版（Microsoft YaHei + 紧凑行距 + 字号），而不是用 pandoc 重新生成一份西方排版风格的 docx。

---

## 📁 目录结构

```
Leetcode/
├── README.md                  # 本文件
├── LICENSE                    # MIT
├── .gitignore                 # 忽略构建产物
├── 文档版/                    # ⭐ 源（Word）
│   └── 晨读背诵内容v0.5.docx
├── PDF版/                     # pdf 输出（由 build.sh 生成）
├── solution/                  # md 拆分产物（由 split-by-problem.py 生成）
│   ├── assets/                # 题解配图
│   ├── 0001.两数之和.md
│   └── ... (按需生成)
└── scripts/
    ├── build.sh               # docx → pdf
    └── split-by-problem.py    # docx → md（带图片抽取）
```

---

## 🛠 工具依赖

| 工具 | 版本要求 | 用途 | 安装 |
|------|----------|------|------|
| **LibreOffice** | ≥ 7.0 | docx → pdf 高保真转换 | `brew install --cask libreoffice` / `apt install libreoffice` |
| **python-docx** | ≥ 1.1 | docx → md 解析 | `pip install python-docx` |

**注意：之前版本用过 pandoc 来 md → docx，已弃用。** pandoc 生成的 docx 字号 / 行距 / 字体全是西方默认值（标题 16pt 粗体大字号，行距 1.0），中文显示效果远不如 Word 手动调好的版式。现在 docx 是唯一源，pdf 直接用 LibreOffice 渲染 —— 100% 保留 Word 排版。

---

## 🚀 典型工作流

### 场景 1：做完一题，想更新 pdf

```bash
# 1. 在 Word / WPS 里编辑 文档版/晨读背诵内容v0.5.docx
# 2. 重新生成 pdf
bash scripts/build.sh
# 3. 重新拆分 md（搜索 / git diff 用）
python3 scripts/split-by-problem.py 文档版/晨读背诵内容v0.5.docx
```

### 场景 2：做完一个里程碑，升级版本号

```bash
# 1. 在 Word 里另存为 文档版/晨读背诵内容v0.6.docx
# 2. 生成 pdf
bash scripts/build.sh 文档版/晨读背诵内容v0.6.docx
# 3. 更新 md
python3 scripts/split-by-problem.py 文档版/晨读背诵内容v0.6.docx
```

### 场景 3：指定输出文件名

```bash
bash scripts/build.sh 文档版/晨读背诵内容v0.5.docx Hot100-打印版
# 输出：PDF版/Hot100-打印版.pdf
```

### 默认行为：不传参数

```bash
bash scripts/build.sh
# 自动取 文档版/ 下最新的 .docx，输出同名 pdf 到 PDF版/
```

---

## 📖 题目索引（按官方分类，共 100 题）

**难度分布：** `简单 20 题` · `中等 68 题` · `困难 12 题`

| 章节 | 题数 | 题号 |
|------|------|------|
| 一、哈希 | 3 | 0001 0049 0128 |
| 二、双指针 | 4 | 0283 0011 0015 0042 |
| 三、滑动窗口 | 4 | 0003 0438 0560 0239 |
| 四、子串 | 2 | 0076 0139 |
| 五、普通数组 | 5 | 0053 0056 0189 0238 0041 |
| 六、矩阵 | 4 | 0073 0054 0048 0240 |
| 七、链表 | 14 | 0160 0206 0234 0141 0142 0021 0002 0019 0024 0025 0138 0148 0023 0146 |
| 八、二叉树 | 15 | 0094 0104 0226 0101 0543 0102 0108 0098 0230 0199 0114 0105 0437 0236 0124 |
| 九、图论 | 4 | 0200 0994 0207 0208 |
| 十、回溯 | 8 | 0046 0078 0017 0039 0022 0079 0131 0051 |
| 11、二分查找 | 6 | 0035 0074 0034 0033 0153 0004 |
| 12、栈 | 5 | 0020 0155 0394 0739 0084 |
| 13、堆 | 3 | 0215 0347 0295 |
| 14、贪心算法 | 4 | 0121 0055 0045 0763 |
| 15、动态规划 | 9 | 0070 0118 0198 0279 0322 0300 0152 0416 0032 |
| 16、多维动态规划 | 5 | 0062 0064 0005 1143 0072 |
| 17、技巧 | 5 | 0136 0169 0075 0031 0287 |

> 题目索引只是导航。**每题的具体内容在 docx 里**，这里不重复展示。

---

## 🔧 排版约定（v0.5 源文件已设置）

| 元素 | 字体 | 字号 | 行距 | 段后 |
|------|------|------|------|------|
| 正文 | Microsoft YaHei | 10.5 pt | 1.5 | 40 twips |
| 章节标题（一、哈希）| Microsoft YaHei | 20 pt 粗 | 1.5 | 60 twips |
| 题号 + 题名 | Microsoft YaHei | 15 pt 粗 | 1.5 | 60 twips |
| 难度【简单/中等/困难】| Microsoft YaHei | 13 pt | 1.5 | 40 twips |
| 学霸笔记 ✎ | Microsoft YaHei | 13 pt 粗体 | 1.6 | 60 twips |

> 改 docx 排版时保持这套数值，新生成的 pdf 才能跟历史版本视觉一致。

---

## 📜 License

[MIT](./LICENSE) · © 2026 gugu

---

