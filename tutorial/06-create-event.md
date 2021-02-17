---
ms.openlocfilehash: facdbb5c42e60e5bb0ee98b06ef68939a6010a67
ms.sourcegitcommit: 24bb4b3df6a035806a58b609e1d8078ac5505fa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50274214"
---
<!-- markdownlint-disable MD002 MD041 -->

<span data-ttu-id="e4554-101">在此部分中，您将添加在用户日历上创建事件的能力。</span><span class="sxs-lookup"><span data-stu-id="e4554-101">In this section you will add the ability to create events on the user's calendar.</span></span>

## <a name="implement-the-api"></a><span data-ttu-id="e4554-102">实现 API</span><span class="sxs-lookup"><span data-stu-id="e4554-102">Implement the API</span></span>

1. <span data-ttu-id="e4554-103">打开 **./src/api/graph.ts** 并添加以下代码，以实施新的事件 API `POST /graph/newevent` () 。</span><span class="sxs-lookup"><span data-stu-id="e4554-103">Open **./src/api/graph.ts** and add the following code to implement a new event API (`POST /graph/newevent`).</span></span>

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="CreateEventSnippet":::

1. <span data-ttu-id="e4554-104">打开 **./src/addin/taskpane.js** 并添加以下函数以调用新的事件 API。</span><span class="sxs-lookup"><span data-stu-id="e4554-104">Open **./src/addin/taskpane.js** and add the following function to call the new event API.</span></span>

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="CreateEventSnippet":::

1. <span data-ttu-id="e4554-105">保存所有更改，重新启动服务器，并刷新 Excel 中的任务窗格 (关闭任何打开的任务窗格，然后重新打开) 。</span><span class="sxs-lookup"><span data-stu-id="e4554-105">Save all of your changes, restart the server, and refresh the task pane in Excel (close any open task panes and re-open).</span></span>

    ![创建事件表单的屏幕截图](images/create-event-ui.png)

1. <span data-ttu-id="e4554-107">填写表单，然后选择"创建 **"。**</span><span class="sxs-lookup"><span data-stu-id="e4554-107">Fill in the form and choose **Create**.</span></span> <span data-ttu-id="e4554-108">验证事件是否添加到用户日历。</span><span class="sxs-lookup"><span data-stu-id="e4554-108">Verify that the event is added to the user's calendar.</span></span>
