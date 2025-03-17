# Copy and paste dimensions in Powerpoint

Select an object, copy its dimensions (width, height, and location on the slide), and apply these dimensions to another object. Especially helpful if you want to align objects across several pages.


![Screen recording showing how to copy dimensions](/assets/screen-recording-copy-dimensions.gif)


## Installation

**Currently only Mac is supported**

* Close Powerpoint.

* Download the [manifest.xml](https://copy-dimensions.vercel.app/manifest.xml) file and place it in the following folder:
 
    `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`
 
    If this folder does not exist, please create it. This process is called [side-loading](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing).

* Open a fresh instance of Powerpoint, go to "Home", then "Add-ins", the add-in should be available under "Developer Add-ins".


## To do 

* Visual indication that clicking on "copy" and "paste" was successful
* Vibe coding in powerpoint: link textbox to LLM, ask it to produce a typescript funection implementing the ask & then directly perform the action
* Rounded boxes change the corner radius when resized. Copy paste the border radius from one box to the other
