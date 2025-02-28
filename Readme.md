# Copy and paste dimensions in Powerpoint

Select one object, copy its dimensions (width, height, and location on the slide), and apply these dimensions to another object. Especially helpful if you want to align objects across several pages.

## How to run

* Download or clone the repository

* Run: 
`npm start`

* A fresh Powerpoint presentation should open with the add-in visible. 


## Troubleshooting

If the add-in is not immediately visible, you can try the following things:

* Check the [folder locations](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing) for "side-loading" add-ins, this is the process that has been used in this setup to make the application visible. For example, on MacOS, the there should be a `something.manifest.xml` file in the folder `/Users/XXXX/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

* Within Powerpoint, go to "Home", then "Add-ins", the add-in should be available under "Developer Add-ins".

