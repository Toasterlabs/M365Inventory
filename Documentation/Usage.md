#Usage
## Introduction
The usage of the tool has been kept as simple as possible intentionally. Below is a step-by-step walkthrough of the tool as it stands on February 2nd, 2019.

## Getting Started
There are 2 ways to start the tool:

### Normal
![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/Start.jpg)

Starts the tool normally.

### Debugging
![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/Start%20-%20Debug.jpg)

Starts the tool with global Synchronized Hashes. This means that $Synchash and $VariableHash will be exposed outside the tools runspace. This can be usefull for debugging purposes, since you can view if all values are loaded correctly.

## Loaded
![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/Loaded.jpg)

This screen will be presented once the tool loads. As you can see, the first thing to happen is prerequisite checks occuring. It is recommended that all checks should produce a green light!

## Settings
![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/settings.png)

The next thing to do is click on the settings tab. This tab must be populated with all the settings for the tool to run corrrectly.

### Browse
![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/Browse.png)

![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/Browse%20-%20Dialog.png)

The browse button will launch a browse dialog so you can select the right output folder. The text field cannot be edited manually.

### Enter Credentials
![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/EnterCredentials.png)

Enter your credentials and hit the validate button.
![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/Validate-Failure.png)

Incorrect credentials will trigger a warning

![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/Validate-Success.png)

While correct credentials will return a success condition and allow you to proceed.

### Complete the remaining fields
![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/CompleteSettings.png)

Now it's time to complete the remaining fields.

## Return home
![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/BackToHome.png)

Click on the home tab. You'll notice that the "Credentials validated" light is now set to green.

## Run the reporting!
Atlast it is time to run the reporting! Click on **File** > **Run** and things will be set in motion...

![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/RunReports.png)

During the run phase, lights will go green or red depending on success or failure of the reporting piece.

![alt text](https://github.com/Toasterlabs/M365Inventory/blob/master/Documentation/Images/Running.png)

1. **Connections** are the first ones who will light up as green. Each represents a succes (green), waiting (orange), or failure (red) condition. Note that at the end of the reporting run, the Exchange Online light will return to Orange, as the sessions gets disconnected. This way the tool can be closed cleanly, and no wait time is required to release the sessions (15 minutes normally)
2. During the reporting run the **activities** field will update with the current action in progress. The field **should** automatically scroll to the end, however I've noticed that it doesn't always do this. I've been able to reproduce the behavior everytime I move or click the interface. If you know how to solve that, for the love of God, please tell me!
3. **Reports** will light up green one by one, not necessarely in the order of which they are listed.
