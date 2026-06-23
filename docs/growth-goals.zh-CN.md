# 公开增长目标

本文定义 Microsoft Excel BI Agent 的公开增长目标。它与维护目标分开：维护目标保证项目安全可靠，增长目标负责让价值更容易被理解、信任和安装。

维护者：**丘彬彬**<br>
微信：**binstudy**<br>
博客：**https://90le.cn**

## 目标

让仓库在访问者进入后的一分钟内建立可信度：访问者应能快速理解这个技能包做什么、什么时候该用、为什么真实 Excel BI 工作簿对普通 Agent 有风险、如何安装、由谁维护，以及哪些验证边界仍然存在。

目标不是泛泛做广告。目标是面向真正处理 Excel BI 工作簿、交付风险和 AI Agent 自动化的人，提高高信任采用率。

## 约束

- 安装说明必须真实。未发布 npm 包前，不宣传 npm 或 npx 命令。
- 英文和中文公开页面继续独立维护。
- 不把客户文件、截图、工作簿报告、私有路径、凭证或运行证据当作营销素材。
- 文本密集型视觉优先使用确定性的仓库资产。只有在确实提升理解且不会产生错误文字时，才使用生成图片。
- 不能暗示结构校验已经证明 Excel COM、VBA、Power Query 刷新或 Power Pivot 运行时行为。
- 网站必须在桌面和移动端保持可读。

## 边界

- 公开增长工作可以优化 README 文案、Pages 布局、社交预览元信息、截图、release 可见性、使用场景表达和广告文案。
- 公开增长工作不应修改 Excel 工作簿处理行为，除非另有产品质量目标支撑。
- 仓库可以展示脱敏样例和工作流形态，但不展示私有工作簿证据。
- 项目可以包含联系方式和维护者署名，但 README 不应变成个人主页。

## 可以做

- 在 README 和 Pages 增加清晰定位、使用场景和信任信息。
- 增加作者/维护者署名和联系方式。
- 增加可复用发布文案、社交文案和广告文案，并指向真实安装路径。
- 使用绝对公开图片 URL 优化 Open Graph/Twitter 元信息。
- 增加校验，防止版本、署名或增长文档在发布时遗漏。
- 让营销语言基于真实技能和验证边界，而不是夸大承诺。

## 不能做

- 添加假的包管理命令、假的 benchmark、假的 logo 或项目不支持的集成。
- 使用私有客户工作簿图片或本地运行 artifact 做宣传。
- 没有 Windows 桌面版 Excel 证据时，承诺 Excel 运行时行为。
- 增加拖慢网站或遮挡安装命令的装饰性重图页面。
- 把长篇中英文内容混排在同一个面向用户的文档里。

## 详细 Goal

| Goal | 价值 | 完成标准 |
| --- | --- | --- |
| 一分钟清晰度 | 新访问者能快速判断项目是否适合自己的 Excel BI 风险。 | README 和 Pages 靠前位置说明受众、使用场景、安装方式和运行时边界。 |
| 信任与署名 | 开源用户需要知道项目由谁维护。 | README、Pages 和元信息展示丘彬彬、微信 `binstudy`、博客 `90le.cn`。 |
| 社交分享 | 链接分享时应渲染正确的预览图和描述。 | HTML 页面使用绝对 Open Graph/Twitter 图片 URL 和 canonical URL。 |
| 广告准备度 | 维护者可以直接复用文案推广仓库。 | 中英文营销文案包提供发布文案、短广告和渠道变体。 |
| 图文丰富但不冒险 | 站点应像真实产品页，但不能编造证据。 | 复用已有确定性资产；生成图片是可选项，不是必需项。 |
| 由校验托底的增长 | 营销优化不能绕过质量门禁。 | 文档校验和 goal coverage 纳入增长文档与署名检查。 |

## 高价值优化 Backlog

| 优先级 | 事项 | 为什么值得做 | 边界 |
| --- | --- | --- | --- |
| P0 | 增加维护者署名和联系方式。 | 提升信任和归属感，运行风险极低。 | 保持克制，不把项目变成个人品牌首页。 |
| P0 | 把社交元信息改为绝对且当前的链接。 | 预览失败会降低分享价值。 | 只使用公开资产，不使用私有图片。 |
| P0 | 增加公开增长目标和广告文案。 | 把模糊营销工作变成可复用、可审阅的契约。 | 文案只能提真实能力和真实安装方式。 |
| P1 | 在 Pages 增加使用场景转化区。 | 帮用户把工作簿痛点映射到技能包价值。 | 安装命令继续可见且移动端可读。 |
| P1 | 在 CI 中校验增长文档。 | 避免 release 之后营销说法漂移。 | CI 仍为结构校验，不依赖桌面版 Excel。 |
| P2 | 后续增加短 demo GIF 或视频。 | 能帮助非技术访问者理解。 | 必须基于脱敏样例，不能包含客户数据。 |

## 必跑公开校验

修改 README、网站、营销文案、release notes、安装文档或校验脚本后运行：

```bash
python tools/validate-skills.py .
python tools/validate_project_docs.py --project-root .
python tools/validate_task_recipes.py --project-root .
python tools/validate_official_docs_index.py --project-root .
python tools/build_artifact_hygiene_report.py --project-root . --require-pass
python tools/build_goal_coverage_report.py --project-root . --require-pass
node tools/install.mjs --check
```
