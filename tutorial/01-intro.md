---
ms.openlocfilehash: 2c323d61632c62c82af0561536656f1d68fe8938
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274233"
---
<!-- markdownlint-disable MD002 MD041 -->

本教程指导你如何构建使用 Microsoft Graph API 检索用户的日历信息的适用于 Excel 的 Office 外接程序。

> [!TIP]
> 如果你更喜欢仅下载已完成的教程，可以下载或克隆 [GitHub 存储库](https://github.com/microsoftgraph/msgraph-training-office-addin)。

## <a name="prerequisites"></a>先决条件

在开始此演示之前，应在开发计算机上安装Node.js[](https://nodejs.org)[和一](https://yarnpkg.com/)个。 如果你没有安装Node.js，请访问上一链接，查看下载选项。

> [!NOTE]
> Windows 用户可能需要安装 Python 和 Visual Studio 生成工具，以支持需要从 C/C++ 编译的 NPM 模块。 Windows Node.js安装程序提供了自动安装这些工具的选项。 或者，您也可以按照以下说明操作 [https://github.com/nodejs/node-gyp#on-windows](https://github.com/nodejs/node-gyp#on-windows) 。

你还应该拥有具有邮箱的个人 Microsoft 帐户Outlook.com或 Microsoft 工作或学校帐户。 如果你没有 Microsoft 帐户，则有两种获取免费帐户的选项：

- 你可以 [注册新的个人 Microsoft 帐户](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1)。
- 你可以 [注册 Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program) ，获取免费的 Office 365 订阅。

> [!NOTE]
> 本教程使用 Node 版本 14.15.0 和一线 1.22.0 编写。 本指南中的步骤可能与其他版本一起运行，但尚未经过测试。

## <a name="feedback"></a>反馈

请在 GitHub 存储库中提供有关本教程 [的任何反馈](https://github.com/microsoftgraph/msgraph-training-office-addin)。
