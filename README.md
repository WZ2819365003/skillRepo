# SkillRepo - 私人定制技能库

私人定制的 Claude/Claude Code 技能库，不定期更新。分为通用技能和专业技能，把个人能力具象化。

## 技能列表

| 技能 | 目录 | 说明 |
|------|------|------|
| **common** | `common/` | 通用文档结构化生成技能。解析参考文件结构，按结构分块生成新内容的 Markdown，调用 docx-cn 转 Word |
| **docx-cn** | `docx-cn/` | 中文 Word 文档生成技能。支持多模板（通用文档/学位论文/期刊论文/技术报告）、公式编排、目录、参考文献、中英混排字体 |
| **state-grid-keyan** | `state-grid-keyan/` | 国家电网/南方电网科技项目可行性研究报告生成技能。按国网/南网标准格式分章节多轮生成完整可研报告 |

## 目录结构

```
skillRepo/
├── README.md
├── LICENSE
├── .gitignore
├── common/                    # 通用文档生成技能
│   ├── .claude/skills/common/SKILL.md
│   └── .claude-plugin/plugin.json
├── docx-cn/                   # 中文Word生成技能
│   ├── .claude/skills/docx-cn/SKILL.md
│   ├── .claude/settings.local.json
│   ├── .claude-plugin/plugin.json
│   ├── imag/
│   ├── test/
│   ├── package.json
│   └── package-lock.json
└── state-grid-keyan/          # 国网可研报告技能
    ├── .claude/skills/state-grid-keyan/SKILL.md
    ├── .claude/settings.local.json
    └── .claude-plugin/plugin.json
```

## 安装使用

### 克隆仓库

```bash
git clone https://github.com/WZ2819365003/skillRepo.git
```

### 安装 docx-cn 依赖

```bash
cd skillRepo/docx-cn
npm install
```

### 安装技能到 Claude Code

将对应技能目录作为 Claude Code 插件安装即可使用。

## 许可证

MIT License
