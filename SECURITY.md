# Security Policy

Microsoft Excel BI Agent is a public repository for Excel BI agent workflows. Security reports often involve workbook data, credentials, provider strings, local paths, or generated diagnostic evidence, so public disclosure can create additional risk.

## Do Not Put These In Public Issues

- Customer workbooks, screenshots, PDFs, or exports. Do not share customer workbooks in public issues.
- Credentials, tokens, connection strings, provider secrets, or local private paths.
- Generated QA reports or runtime evidence that has not been deliberately sanitized.
- Private workbook formulas, VBA modules, Power Query sources, DAX measures, or CUBE references that reveal business data.

## Report Privately

If the report includes sensitive details, do not open a public issue. Use GitHub private vulnerability reporting when it is available for this repository, or contact the maintainer directly:

- Maintainer: **Qiu Binbin (丘彬彬)**
- WeChat: **binstudy**
- Blog: **https://90le.cn**

If you can reproduce the issue with a sanitized fixture, open a public issue using the appropriate issue form and include only the sanitized reproduction.

## Supported Scope

Security-relevant reports include:

- accidental inclusion of customer data or local paths in repository files;
- install instructions that could lead users to run unsupported or unsafe commands;
- helper scripts that mishandle credentials, provider strings, generated reports, or workbook artifacts;
- validation gaps that allow customer artifacts into the public package.

## Runtime Boundary

Structural validation on macOS/Linux does not prove Excel COM, VBA, Power Query refresh, or Power Pivot runtime behavior. Runtime proof must come from Windows desktop Excel and must not include private workbook evidence in public issues or pull requests.
