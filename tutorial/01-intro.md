---
ms.openlocfilehash: 2c323d61632c62c82af0561536656f1d68fe8938
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274233"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="ff93f-101">本教程指导你如何构建使用 Microsoft Graph API 检索用户的日历信息的适用于 Excel 的 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="ff93f-101">This tutorial teaches you how to build an Office Add-in for Excel that uses the Microsoft Graph API to retrieve calendar information for a user.</span></span>

> [!TIP]
> <span data-ttu-id="ff93f-102">如果你更喜欢仅下载已完成的教程，可以下载或克隆 [GitHub 存储库](https://github.com/microsoftgraph/msgraph-training-office-addin)。</span><span class="sxs-lookup"><span data-stu-id="ff93f-102">If you prefer to just download the completed tutorial, you can download or clone the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-office-addin).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ff93f-103">先决条件</span><span class="sxs-lookup"><span data-stu-id="ff93f-103">Prerequisites</span></span>

<span data-ttu-id="ff93f-104">在开始此演示之前，应在开发计算机上安装Node.js[](https://nodejs.org)[和一](https://yarnpkg.com/)个。</span><span class="sxs-lookup"><span data-stu-id="ff93f-104">Before you start this demo, you should have [Node.js](https://nodejs.org) and [Yarn](https://yarnpkg.com/) installed on your development machine.</span></span> <span data-ttu-id="ff93f-105">如果你没有安装Node.js，请访问上一链接，查看下载选项。</span><span class="sxs-lookup"><span data-stu-id="ff93f-105">If you do not have Node.js or Yarn, visit the previous link for download options.</span></span>

> [!NOTE]
> <span data-ttu-id="ff93f-106">Windows 用户可能需要安装 Python 和 Visual Studio 生成工具，以支持需要从 C/C++ 编译的 NPM 模块。</span><span class="sxs-lookup"><span data-stu-id="ff93f-106">Windows users may need to install Python and Visual Studio Build Tools to support NPM modules that need to be compiled from C/C++.</span></span> <span data-ttu-id="ff93f-107">Windows Node.js安装程序提供了自动安装这些工具的选项。</span><span class="sxs-lookup"><span data-stu-id="ff93f-107">The Node.js installer on Windows gives an option to automatically install these tools.</span></span> <span data-ttu-id="ff93f-108">或者，您也可以按照以下说明操作 [https://github.com/nodejs/node-gyp#on-windows](https://github.com/nodejs/node-gyp#on-windows) 。</span><span class="sxs-lookup"><span data-stu-id="ff93f-108">Alternatively, you can follow instructions at [https://github.com/nodejs/node-gyp#on-windows](https://github.com/nodejs/node-gyp#on-windows).</span></span>

<span data-ttu-id="ff93f-109">你还应该拥有具有邮箱的个人 Microsoft 帐户Outlook.com或 Microsoft 工作或学校帐户。</span><span class="sxs-lookup"><span data-stu-id="ff93f-109">You should also have either a personal Microsoft account with a mailbox on Outlook.com, or a Microsoft work or school account.</span></span> <span data-ttu-id="ff93f-110">如果你没有 Microsoft 帐户，则有两种获取免费帐户的选项：</span><span class="sxs-lookup"><span data-stu-id="ff93f-110">If you don't have a Microsoft account, there are a couple of options to get a free account:</span></span>

- <span data-ttu-id="ff93f-111">你可以 [注册新的个人 Microsoft 帐户](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1)。</span><span class="sxs-lookup"><span data-stu-id="ff93f-111">You can [sign up for a new personal Microsoft account](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1).</span></span>
- <span data-ttu-id="ff93f-112">你可以 [注册 Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program) ，获取免费的 Office 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="ff93f-112">You can [sign up for the Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get a free Office 365 subscription.</span></span>

> [!NOTE]
> <span data-ttu-id="ff93f-113">本教程使用 Node 版本 14.15.0 和一线 1.22.0 编写。</span><span class="sxs-lookup"><span data-stu-id="ff93f-113">This tutorial was written with Node version 14.15.0 and Yarn version 1.22.0.</span></span> <span data-ttu-id="ff93f-114">本指南中的步骤可能与其他版本一起运行，但尚未经过测试。</span><span class="sxs-lookup"><span data-stu-id="ff93f-114">The steps in this guide may work with other versions, but that has not been tested.</span></span>

## <a name="feedback"></a><span data-ttu-id="ff93f-115">反馈</span><span class="sxs-lookup"><span data-stu-id="ff93f-115">Feedback</span></span>

<span data-ttu-id="ff93f-116">请在 GitHub 存储库中提供有关本教程 [的任何反馈](https://github.com/microsoftgraph/msgraph-training-office-addin)。</span><span class="sxs-lookup"><span data-stu-id="ff93f-116">Please provide any feedback on this tutorial in the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-office-addin).</span></span>
