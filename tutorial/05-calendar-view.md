---
ms.openlocfilehash: ade142f1518d9bd56fa1472889721a4c8995bc3c
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274215"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="93a0d-101">在此练习中，你将将 Microsoft Graph 合并到应用程序中。</span><span class="sxs-lookup"><span data-stu-id="93a0d-101">In this exercise you will incorporate Microsoft Graph into the application.</span></span> <span data-ttu-id="93a0d-102">对于此应用程序，你将使用 [microsoft-graph-client](https://github.com/microsoftgraph/msgraph-sdk-javascript) 库调用 Microsoft Graph。</span><span class="sxs-lookup"><span data-stu-id="93a0d-102">For this application, you will use the [microsoft-graph-client](https://github.com/microsoftgraph/msgraph-sdk-javascript) library to make calls to Microsoft Graph.</span></span>

## <a name="get-calendar-events-from-outlook"></a><span data-ttu-id="93a0d-103">从 Outlook 获取日历事件</span><span class="sxs-lookup"><span data-stu-id="93a0d-103">Get calendar events from Outlook</span></span>

<span data-ttu-id="93a0d-104">首先添加 API， [从用户的](https://docs.microsoft.com/graph/api/user-list-calendarview) 日历获取日历视图。</span><span class="sxs-lookup"><span data-stu-id="93a0d-104">Start by adding an API to get a [calendar view](https://docs.microsoft.com/graph/api/user-list-calendarview) from the user's calendar.</span></span>

1. <span data-ttu-id="93a0d-105">打开 **./src/api/graph.ts，** 将以下语句 `import` 添加到文件顶部。</span><span class="sxs-lookup"><span data-stu-id="93a0d-105">Open **./src/api/graph.ts** and add the following `import` statements to the top of the file.</span></span>

    ```typescript
    import { zonedTimeToUtc } from 'date-fns-tz';
    import { findOneIana } from 'windows-iana';
    import * as graph from '@microsoft/microsoft-graph-client';
    import { Event, MailboxSettings } from 'microsoft-graph';
    import 'isomorphic-fetch';
    import { getTokenOnBehalfOf } from './auth';
    ```

1. <span data-ttu-id="93a0d-106">添加以下函数以初始化 Microsoft Graph SDK 并返回 **客户端**。</span><span class="sxs-lookup"><span data-stu-id="93a0d-106">Add the following function to initialize the Microsoft Graph SDK and return a **Client**.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetClientSnippet":::

1. <span data-ttu-id="93a0d-107">添加以下函数，从用户的邮箱设置获取用户的时区，并将该值转换为 IANA 时区标识符。</span><span class="sxs-lookup"><span data-stu-id="93a0d-107">Add the following function to get the user's time zone from their mailbox settings, and to convert that value to an IANA time zone identifier.</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetTimeZonesSnippet":::

1. <span data-ttu-id="93a0d-108">将以下函数 (`const graphRouter = Router();` 行) 实现 API 终结点 `GET /graph/calendarview` () 。</span><span class="sxs-lookup"><span data-stu-id="93a0d-108">Add the following function (below the `const graphRouter = Router();` line) to implement an API endpoint (`GET /graph/calendarview`).</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="GetCalendarViewSnippet":::

    <span data-ttu-id="93a0d-109">考虑此代码执行什么功能。</span><span class="sxs-lookup"><span data-stu-id="93a0d-109">Consider what this code does.</span></span>

    - <span data-ttu-id="93a0d-110">它获取用户的时区，并使用该时区将请求的日历视图的开始和结尾转换为 UTC 值。</span><span class="sxs-lookup"><span data-stu-id="93a0d-110">It gets the user's time zone and uses that to convert the start and end of the requested calendar view into UTC values.</span></span>
    - <span data-ttu-id="93a0d-111">它会对 `GET` Graph `/me/calendarview` API 终结点进行调用。</span><span class="sxs-lookup"><span data-stu-id="93a0d-111">It does a `GET` to the `/me/calendarview` Graph API endpoint.</span></span>
        - <span data-ttu-id="93a0d-112">它使用该函数来设置标头，从而根据用户的时区调整返回事件的开始时间和 `header` `Prefer: outlook.timezone` 结束时间。</span><span class="sxs-lookup"><span data-stu-id="93a0d-112">It uses the `header` function to set the `Prefer: outlook.timezone` header, causing the start and end times of the returned events to be adjusted to the user's time zone.</span></span>
        - <span data-ttu-id="93a0d-113">它使用该 `query` 函数添加 and `startDateTime` `endDateTime` 参数，并设置日历视图的开始和结束。</span><span class="sxs-lookup"><span data-stu-id="93a0d-113">It uses the `query` function to add the `startDateTime` and `endDateTime` parameters, setting the start and end of the calendar view.</span></span>
        - <span data-ttu-id="93a0d-114">它使用该 `select` 函数仅请求外接程序使用的字段。</span><span class="sxs-lookup"><span data-stu-id="93a0d-114">It uses the `select` function to request only the fields used by the add-in.</span></span>
        - <span data-ttu-id="93a0d-115">它使用 `orderby` 函数按开始时间对结果进行排序。</span><span class="sxs-lookup"><span data-stu-id="93a0d-115">It uses the `orderby` function to sort the results by the start time.</span></span>
        - <span data-ttu-id="93a0d-116">它使用该 `top` 函数将单个请求中的结果限制为 25。</span><span class="sxs-lookup"><span data-stu-id="93a0d-116">It uses the `top` function to limit the results in a single request to 25.</span></span>
    - <span data-ttu-id="93a0d-117">它使用 **PageIteratorCallback** 对象 [](https://docs.microsoft.com/graph/sdks/paging)来对结果进行引用，并可在更多结果页面可用时提出其他请求。</span><span class="sxs-lookup"><span data-stu-id="93a0d-117">It uses a **PageIteratorCallback** object to [iterate through the results](https://docs.microsoft.com/graph/sdks/paging) and to make additional requests if more pages of results are available.</span></span>

## <a name="update-the-ui"></a><span data-ttu-id="93a0d-118">更新 UI</span><span class="sxs-lookup"><span data-stu-id="93a0d-118">Update the UI</span></span>

<span data-ttu-id="93a0d-119">现在，让我们更新任务窗格，以允许用户指定日历视图的开始日期和结束日期。</span><span class="sxs-lookup"><span data-stu-id="93a0d-119">Now let's update the task pane to allow the user to specify a start and end date for the calendar view.</span></span>

1. <span data-ttu-id="93a0d-120">打开 **./src/addin/taskpane.js，** 将 `showMainUi` 现有函数替换为以下内容。</span><span class="sxs-lookup"><span data-stu-id="93a0d-120">Open **./src/addin/taskpane.js** and replace the existing `showMainUi` function with the following.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="MainUiSnippet":::

    <span data-ttu-id="93a0d-121">此代码添加一个简单的表单，以便用户可以指定开始日期和结束日期。</span><span class="sxs-lookup"><span data-stu-id="93a0d-121">This code adds a simple form so the user can specify a start and end date.</span></span> <span data-ttu-id="93a0d-122">它还实现用于创建新事件的第二个表单。</span><span class="sxs-lookup"><span data-stu-id="93a0d-122">It also implements a second form for creating a new event.</span></span> <span data-ttu-id="93a0d-123">该表单目前未执行任何功能，您将在下一节中实现该功能。</span><span class="sxs-lookup"><span data-stu-id="93a0d-123">That form doesn't do anything for now, you'll implement that feature in the next section.</span></span>

1. <span data-ttu-id="93a0d-124">将以下代码添加到文件以在活动工作表中创建一个表，其中包含从日历视图中检索到的事件。</span><span class="sxs-lookup"><span data-stu-id="93a0d-124">Add the following code to the file to create a table in the active worksheet containing the events retrieved from the calendar view.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="WriteToSheetSnippet":::

1. <span data-ttu-id="93a0d-125">添加以下函数以调用日历视图 API。</span><span class="sxs-lookup"><span data-stu-id="93a0d-125">Add the following function to call the calendar view API.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="GetCalendarSnippet":::

1. <span data-ttu-id="93a0d-126">保存所有更改，重新启动服务器，并刷新 Excel 中的任务窗格 (关闭任何打开的任务窗格，然后重新打开) 。</span><span class="sxs-lookup"><span data-stu-id="93a0d-126">Save all of your changes, restart the server, and refresh the task pane in Excel (close any open task panes and re-open).</span></span>

    ![导入表单的屏幕截图](images/get-calendar-view-ui.png)

1. <span data-ttu-id="93a0d-128">选择开始日期和结束日期，然后选择"**导入"。**</span><span class="sxs-lookup"><span data-stu-id="93a0d-128">Choose start and end dates and choose **Import**.</span></span>

    ![事件表的屏幕截图](images/calendar-view-table.png)
