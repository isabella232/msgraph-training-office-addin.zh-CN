---
ms.openlocfilehash: a336095a238eefeae22dac86d29e3140e8d94bf1
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274220"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="3b6dc-101">在此练习中，你将使用 Azure Active Directory 管理中心创建新的 Azure AD Web 应用程序注册。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-101">In this exercise, you will create a new Azure AD web application registration using the Azure Active Directory admin center.</span></span>

1. <span data-ttu-id="3b6dc-102">打开浏览器，并转到 [Azure Active Directory 管理中心](https://aad.portal.azure.com)。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-102">Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com).</span></span> <span data-ttu-id="3b6dc-103">使用 **个人帐户**（亦称为“Microsoft 帐户”）或 **工作或学校帐户** 登录。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-103">Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.</span></span>

1. <span data-ttu-id="3b6dc-104">选择左侧导航栏中的“**Azure Active Directory**”，再选择“**管理**”下的“**应用注册**”。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-104">Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.</span></span>

    ![<span data-ttu-id="3b6dc-105">应用注册的屏幕截图</span><span class="sxs-lookup"><span data-stu-id="3b6dc-105">A screenshot of the App registrations</span></span> ](images/app-registrations.png)

1. <span data-ttu-id="3b6dc-106">选择“新注册”。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-106">Select **New registration**.</span></span> <span data-ttu-id="3b6dc-107">在“注册应用”页上，按如下方式设置值。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-107">On the **Register an application** page, set the values as follows.</span></span>

    - <span data-ttu-id="3b6dc-108">将“名称”设置为“`Office Add-in Graph Tutorial`”。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-108">Set **Name** to `Office Add-in Graph Tutorial`.</span></span>
    - <span data-ttu-id="3b6dc-109">将“受支持的帐户类型”设置为“任何组织目录中的帐户和个人 Microsoft 帐户”。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-109">Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.</span></span>
    - <span data-ttu-id="3b6dc-110">在“重定向 URI”下，将第一个下拉列表设置为“`Single-page application (SPA)`”，并将值设置为“`https://localhost:3000/consent.html`”。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-110">Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `https://localhost:3000/consent.html`.</span></span>

    !["注册应用程序"页的屏幕截图](images/register-an-app.png)

1. <span data-ttu-id="3b6dc-112">选择“**注册**”。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-112">Select **Register**.</span></span> <span data-ttu-id="3b6dc-113">在 **"Office 外接程序图形** 教程"页上，复制应用程序应用程序 (客户端) **ID** 并保存它，下一步将需要它。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-113">On the **Office Add-in Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.</span></span>

    ![新应用注册的应用程序 ID 屏幕截图](images/application-id.png)

1. <span data-ttu-id="3b6dc-115">选择“管理”下的“证书和密码”。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-115">Select **Certificates & secrets** under **Manage**.</span></span> <span data-ttu-id="3b6dc-116">选择“新客户端密码”按钮。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-116">Select the **New client secret** button.</span></span> <span data-ttu-id="3b6dc-117">在 Description 中 **输入** 值，然后选择"过期"选项之 **一，然后选择**"**添加"。**</span><span class="sxs-lookup"><span data-stu-id="3b6dc-117">Enter a value in **Description** and select one of the options for **Expires** and select **Add**.</span></span>

1. <span data-ttu-id="3b6dc-118">离开此页前，先复制客户端密码值。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-118">Copy the client secret value before you leave this page.</span></span> <span data-ttu-id="3b6dc-119">将在下一步中用到它。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-119">You will need it in the next step.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="3b6dc-120">此客户端密码不会再次显示，所以请务必现在就复制它。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-120">This client secret is never shown again, so make sure you copy it now.</span></span>

1. <span data-ttu-id="3b6dc-121">在 **"管理"下\*\*\*\*选择** API 权限，然后选择 **"添加权限"。**</span><span class="sxs-lookup"><span data-stu-id="3b6dc-121">Select **API permissions** under **Manage**, then select **Add a permission**.</span></span>

1. <span data-ttu-id="3b6dc-122">选择 **Microsoft Graph，** 然后选择 **委派权限**。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-122">Select **Microsoft Graph**, then **Delegated permissions**.</span></span>

1. <span data-ttu-id="3b6dc-123">选择以下权限，然后选择 **"添加权限"。**</span><span class="sxs-lookup"><span data-stu-id="3b6dc-123">Select the following permissions, then select **Add permissions**.</span></span>

    - <span data-ttu-id="3b6dc-124">**offline_access** - 这将允许应用在令牌过期时刷新访问令牌。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-124">**offline_access** - this will allow the app to refresh access tokens when they expire.</span></span>
    - <span data-ttu-id="3b6dc-125">**Calendars.ReadWrite** - 这将允许应用读取和写入用户的日历。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-125">**Calendars.ReadWrite** - this will allow the app to read and write to the user's calendar.</span></span>
    - <span data-ttu-id="3b6dc-126">**MailboxSettings.Read** - 这将允许应用从用户的邮箱设置获取用户的时区。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-126">**MailboxSettings.Read** - this will allow the app to get the user's time zone from their mailbox settings.</span></span>

    ![已配置权限的屏幕截图](images/configured-permissions.png)

## <a name="configure-office-add-in-single-sign-on"></a><span data-ttu-id="3b6dc-128">配置 Office 外接程序单一登录</span><span class="sxs-lookup"><span data-stu-id="3b6dc-128">Configure Office Add-in single sign-on</span></span>

<span data-ttu-id="3b6dc-129">在此部分中，你将更新应用注册以支持 Office 外接程序单一登录[ (SSO) 。 ](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)</span><span class="sxs-lookup"><span data-stu-id="3b6dc-129">In this section you'll update the app registration to support [Office Add-in single sign-on (SSO)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).</span></span>

1. <span data-ttu-id="3b6dc-130">选择 **"公开 API"。**</span><span class="sxs-lookup"><span data-stu-id="3b6dc-130">Select **Expose an API**.</span></span> <span data-ttu-id="3b6dc-131">在此 **API 定义的** 范围部分中，选择 **"添加范围"。**</span><span class="sxs-lookup"><span data-stu-id="3b6dc-131">In the **Scopes defined by this API** section, select **Add a scope**.</span></span> <span data-ttu-id="3b6dc-132">当系统提示设置应用程序 **ID URI** 时，将值设置为 `api://localhost:3000/YOUR_APP_ID_HERE` ， `YOUR_APP_ID_HERE` 以应用程序 ID 替换。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-132">When prompted to set an **Application ID URI**, set the value to `api://localhost:3000/YOUR_APP_ID_HERE`, replacing `YOUR_APP_ID_HERE` with the application ID.</span></span> <span data-ttu-id="3b6dc-133">选择 **"保存"并继续**。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-133">Choose **Save and continue**.</span></span>

1. <span data-ttu-id="3b6dc-134">按如下所示填写字段，然后选择"**添加范围"。**</span><span class="sxs-lookup"><span data-stu-id="3b6dc-134">Fill in the fields as follows and select **Add scope**.</span></span>

    - <span data-ttu-id="3b6dc-135">**范围名称：**`access_as_user`</span><span class="sxs-lookup"><span data-stu-id="3b6dc-135">**Scope name:** `access_as_user`</span></span>
    - <span data-ttu-id="3b6dc-136">**谁可以同意？：管理员和用户**</span><span class="sxs-lookup"><span data-stu-id="3b6dc-136">**Who can consent?: Admins and users**</span></span>
    - <span data-ttu-id="3b6dc-137">**管理员同意显示名称：**`Access the app as the user`</span><span class="sxs-lookup"><span data-stu-id="3b6dc-137">**Admin consent display name:** `Access the app as the user`</span></span>
    - <span data-ttu-id="3b6dc-138">**管理员同意说明：**`Allows Office Add-ins to call the app's web APIs as the current user.`</span><span class="sxs-lookup"><span data-stu-id="3b6dc-138">**Admin consent description:** `Allows Office Add-ins to call the app's web APIs as the current user.`</span></span>
    - <span data-ttu-id="3b6dc-139">**用户同意显示名称：**`Access the app as you`</span><span class="sxs-lookup"><span data-stu-id="3b6dc-139">**User consent display name:** `Access the app as you`</span></span>
    - <span data-ttu-id="3b6dc-140">**用户同意说明：**`Allows Office Add-ins to call the app's web APIs as you.`</span><span class="sxs-lookup"><span data-stu-id="3b6dc-140">**User consent description:** `Allows Office Add-ins to call the app's web APIs as you.`</span></span>
    - <span data-ttu-id="3b6dc-141">**状态：已启用**</span><span class="sxs-lookup"><span data-stu-id="3b6dc-141">**State: Enabled**</span></span>

    !["添加范围"表单的屏幕截图](images/add-scope.png)

1. <span data-ttu-id="3b6dc-143">在"**授权客户端应用程序**"部分，选择 **"添加客户端应用程序"。**</span><span class="sxs-lookup"><span data-stu-id="3b6dc-143">In the **Authorized client applications** section, select **Add a client application**.</span></span> <span data-ttu-id="3b6dc-144">从以下列表中输入客户端 ID，在"授权范围"下启用范围，然后选择"**添加应用程序"。**</span><span class="sxs-lookup"><span data-stu-id="3b6dc-144">Enter a client ID from the following list, enable the scope under **Authorized scopes**, and select **Add application**.</span></span> <span data-ttu-id="3b6dc-145">对列表中的每个客户端 ID 重复此过程。</span><span class="sxs-lookup"><span data-stu-id="3b6dc-145">Repeat this process for each of the client IDs in the list.</span></span>

    - <span data-ttu-id="3b6dc-146">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="3b6dc-146">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="3b6dc-147">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="3b6dc-147">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="3b6dc-148">`57fb890c-0dab-4253-a5e0-7188c88b2bb4`（Office 网页版）</span><span class="sxs-lookup"><span data-stu-id="3b6dc-148">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="3b6dc-149">`08e18876-6177-487e-b8b5-cf950c1e598c`（Office 网页版）</span><span class="sxs-lookup"><span data-stu-id="3b6dc-149">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)</span></span>
