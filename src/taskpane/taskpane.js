/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

function getCategories() {
  Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const categories = asyncResult.value;
      if (categories && categories.length > 0) {
        console.log("Categories assigned to this item:");
        console.log(JSON.stringify(categories));
      } else {
        console.log("There are no categories assigned to this item.");
      }
    } else {
      console.error(asyncResult.error);
    }
  });
}

function addCategories() {
  // Note: In order for you to successfully add a category,
  // it must be in the mailbox categories master list.

  Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const masterCategories = asyncResult.value;
      if (masterCategories && masterCategories.length > 0) {
        // Grab the first category from the master list.
        const categoryToAdd = [masterCategories[0].displayName];
        Office.context.mailbox.item.categories.addAsync(categoryToAdd, function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`Successfully assigned category '${categoryToAdd}' to item.`);
          } else {
            console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
          }
        });
      } else {
        console.log(
          "There are no categories in the master list on this mailbox. You can add categories using Office.context.mailbox.masterCategories.addAsync."
        );
      }
    } else {
      console.error(asyncResult.error);
    }
  });
}

function removeCategories() {
  Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const categories = asyncResult.value;
      if (categories && categories.length > 0) {
        // Grab the first category assigned to this item.
        const categoryToRemove = [categories[0].displayName];
        Office.context.mailbox.item.categories.removeAsync(categoryToRemove, function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`Successfully unassigned category '${categoryToRemove}' from this item.`);
          } else {
            console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
          }
        });
      } else {
        console.log("There are no categories assigned to this item.");
      }
    } else {
      console.error(asyncResult.error);
    }
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    document.getElementById("get-categories").onclick = getCategories;
    document.getElementById("add-categories").onclick = addCategories;
    document.getElementById("remove-categories").onclick = removeCategories;

  }
});

let counter = 1;

export async function run() {
  /**
   * Insert your Outlook code here
   */
  // Get a reference to the current message
  const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject + counter++;

  // Office.context.mailbox.item.body.getAsync(
  //   "text",
  //   { asyncContext: "this is passed to the callback" },
  //   function callback(result) {
  //     document.getElementById("item-body").innerHTML = "<b>Body:</b><br/>" + result;
  //   }
  // );

  Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const masterCategories = asyncResult.value;
      if (masterCategories && masterCategories.length > 0) {
        document.getElementById("item-body").innerHTML = "<b>categories:</b><br/>" + masterCategories[0].displayName;
        // Grab the first category from the master list.
        const categoryToAdd = [masterCategories[0].displayName];
        Office.context.mailbox.item.categories.addAsync(categoryToAdd, function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`Successfully assigned category '${categoryToAdd}' to item.`);
          } else {
            console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
          }
        });
      } else {
        console.log(
          "There are no categories in the master list on this mailbox. You can add categories using Office.context.mailbox.masterCategories.addAsync."
        );
      }
    } else {
      console.error(asyncResult.error);
    }
  });
}
