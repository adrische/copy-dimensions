/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// Store dimensions as a simple object instead
let storedDimensions: { left: number; top: number; width: number; height: number } | null = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Update click handlers for new buttons
    document.getElementById("copy-dims").onclick = copyDimensions;
    document.getElementById("paste-dims").onclick = pasteDimensions;
  }
});

export async function copyDimensions() {
  try {
    // Get the current slide context
    await PowerPoint.run(async (context) => {
      // Get the selected shape
      const shapes = context.presentation.getSelectedShapes();
      context.load(shapes, 'items');
      await context.sync();

      if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        storedDimensions = {
          left: shape.left,
          top: shape.top,
          width: shape.width,
          height: shape.height
        };
        console.log("Dimensions copied successfully", storedDimensions);
      } else {
        console.error("No shape selected");
      }
    });
  } catch (error) {
    console.error("Error copying dimensions:", error);
  }
}

export async function pasteDimensions() {
  try {
    if (!storedDimensions) {
      console.error("No dimensions stored");
      return;
    }

    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      context.load(shapes, 'items');
      await context.sync();

      if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.left = storedDimensions.left;
        shape.top = storedDimensions.top;
        shape.width = storedDimensions.width;
        shape.height = storedDimensions.height;
        await context.sync();
        console.log("Dimensions pasted successfully");
      } else {
        console.error("No shape selected");
      }
    });
  } catch (error) {
    console.error("Error pasting dimensions:", error);
  }
}
