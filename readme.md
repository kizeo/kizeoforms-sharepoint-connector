# SharePoint connector

## The connector

This connector has been created to give to our customers the opportunity of importing/exporting their data with SharePoint.

## Usage

You can run the project with Visual Studio, it's a powerful IDE when it's about coding in C#.

If you want to try out and launch the application, we recommand you to first start the Wizard's project.
It will create a configuration file, aimed to be ran by the connector. Thus, you can choose with the Wizard which forms you want to export, which list on SharePoint you want to import so on and so forth.

### Step 1

You'll be asked some information about your Kizeo Forms' account, as well as some SharePoint's credentials.
If you need any enlightenment about how to get a hand on those credentials, feel free to follow our tutorial.

### Step 2

The next step consists in choosing in which SharePoint's list will updated from a given Kizeo Forms' form. You can go further and decide to which SharePoint's list column a Kizeo Forms' form field is linked.

### Step 3

You can now decide to upload fies to SharePoint docoment library from Kizeo Forms. You'll have the choice among some file formats.

### Step 4

This step's aim is to add the Kizeo Forms' lists you want to update from SharePoint lists.

### Step 5

The final step is about scheduling the uploads from Kizeo Forms to SharePoint. The period be Disabled, weekly, Daily or both. You can now save and finish.

### Start the connector

You can now start the connector. It's a CLI based application where you can watch how your data are transfered. If you let the connector work, a refresh will be performed every 5 minutes.

## Developpment

For the developpment, the code is divided into two parts as well: the connecteur and the wizard.
