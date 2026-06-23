# 营销文案包

维护者：**丘彬彬**<br>
微信：**binstudy**<br>
博客：**https://90le.cn**

用于在 GitHub 动态、社交平台、Newsletter 或内部工具目录中介绍 Microsoft Excel BI Agent。除非项目真的发布了新的分发渠道，否则安装命令必须保持原样。

## 定位

Microsoft Excel BI Agent 是一个开源技能包，用于让 AI Agent 更安全地处理真实 Excel BI 工作簿。它为 VBA、Power Query M、Power Pivot DAX、MDX/CUBE 公式、ADO/SQL、工作簿 QA、纯净交付物、Office 环境诊断、报表搭建、语义模型审阅和脱敏测试样例提供明确工作流。

## 短标语

- 让 AI Agent 更可靠地处理真实 Excel BI 工作簿。
- 面向 Codex、Claude、OpenCode 等 Agent 的 Excel BI 技能包。
- 为 VBA、Power Query、DAX、CUBE 公式和 Excel 交付物提供更安全的 Agent 工作流。
- 在 Agent 修改文件之前，先把工作簿风险显性化。

## 发布文案

Microsoft Excel BI Agent 是一个开源 Excel BI 技能包，面向需要处理真实 Microsoft Excel 工作簿的 AI Agent。

它覆盖 VBA、Power Query M、Power Pivot DAX、MDX/CUBE 公式、ADO/SQL、工作簿 QA、纯净交付物、Office 环境诊断、报表搭建、语义模型审阅和脱敏测试样例。

Codex 插件市场安装：

```bash
codex plugin marketplace add 90le/microsoft-excel-bi-agent
codex plugin add microsoft-excel-bi-agent-pack@microsoft-excel-bi-agent
```

本地安装：

```bash
node tools/install.mjs
```

Excel COM、VBA 执行、Power Query 刷新和 Power Pivot 运行时行为仍然需要 Windows 桌面版 Excel 证明。macOS 和 Linux 支持结构校验和 OpenXML 检查。

## 广告方向

| 方向 | 标题 | 正文 | CTA |
| --- | --- | --- | --- |
| 安装信任 | 不再让 Agent 猜该怎么改 Excel BI 文件。 | 用一个技能包把 VBA、Power Query、DAX、CUBE 公式、QA 和纯净交付纳入明确工作流。 | 从 GitHub 安装 |
| 交付风险 | 工作簿离开本机前，先把风险看清楚。 | 用可复用 Agent 工作流审查公式、隐藏表、外链、VBA、查询和交付副本。 | 查看工作流 |
| 跨 Agent 复用 | 一份 Excel BI 技能源，多种 Agent 表面复用。 | 从同一份技能包同步到 Codex、Claude、OpenCode 风格目录，不编造新命令。 | 查看安装说明 |
| 运行时诚实 | 结构校验有价值，但不是 Excel 运行时证明。 | 项目明确区分 Windows Excel COM、VBA、Power Query 刷新和 Power Pivot 边界。 | 阅读边界 |

## 渠道变体

GitHub 仓库描述：

> 面向 AI Agent 的开源 Excel BI 技能包：覆盖 VBA、Power Query M、Power Pivot DAX、MDX/CUBE、ADO/SQL、工作簿 QA、纯净交付、Office 诊断、报表和脱敏样例。

Newsletter 简介：

> Microsoft Excel BI Agent 为代码 Agent 提供更安全的真实 Excel BI 工作簿操作模型：识别工作簿表面，选择正确 BI 层，小范围修改，并用明确运行时边界验证结果。

短社交文案：

> 新版本：Microsoft Excel BI Agent 帮助 Codex、Claude、OpenCode 等 Agent 更安全地处理 Excel VBA、Power Query、DAX、CUBE 公式、ADO/SQL、QA 和纯净工作簿交付。

## 不要这样宣传

- 未发布 npm 或 npx 包前，不要宣传相关安装命令。
- 未审阅具体工作簿前，不要承诺该工作簿业务准确。
- 不要把 macOS/Linux 结构校验说成 Excel COM、VBA、Power Query 刷新或 Power Pivot 运行时证明。
- 不要使用客户工作簿截图或生成的 QA 报告作为宣传证据。
