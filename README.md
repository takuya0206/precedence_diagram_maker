

This is an add-on for Google apps script. You can automatically create a network chart in PDM (*Precedence Diagram Method) which is suitable to measure total duration of projects. English & Japanese are available.

![f:id:takuya0206:20180318011244g:plain](https://cdn-ak.f.st-hatena.com/images/fotolife/t/takuya0206/20180318/20180318011244.gif)

### Installation

[Precedence Diagram Maker - Google Sheets add-on](https://gsuite.google.com/marketplace/app/precedence_diagram_maker/1091965070131)

Access the above URL, log in your Google account, and click "Install."

### What You Can Do

* Automatically create a network chart in PDM
* Automatically calculate total duration of projects
* Show critical path in red color

### Specification

#### Add-on Menu

Item                      | Action
------------------------- | -------------------
Create Precedence Diagram | Create a list sheet
Show Sidebar              | Show a sidebar

##### Sidebar

Item                        | Action
--------------------------- | ----------------------------------------------------------------------------------------------------
Project Name form（optional） | Include project name in a network chart
Done                        | Create a diagram sheet and automatically create a network chart based on information in the list sheet

#### list sheet

![20180318134335](http://img.f.hatena.ne.jp/images/fotolife/t/takuya0206/20180318/20180318134335.png)

Item          | Note
------------- | -----------------------------------------------------
Activity List | Enter a title of activities （ID is automatically assigned）
Duration      | Enter integer
Precedent ID  | Enter ID of precedent activities. Write 0 if it is the first activity.<br />When there are more than one precedent activity, enter them in comma separated style. e.g. 1,2.
Relationship  | Enter relationship of each activity by choosing one of FS (Finish to Start), SS (Start to Start), SF (Start to Finish), and FF (Finish to Finish).<br />When there are more than one precedent activity, enter them in comma separated style. e.g. FS,SS.
Lead / Lag    | Enter integer. When you enter lead time, use negative number. When you enter lag time, use positive number.<br />When there are more than one precedent activity, enter them in comma separated style. e.g. -10,0.

#### diagramシート

![20180318134338](http://img.f.hatena.ne.jp/images/fotolife/t/takuya0206/20180318/20180318134338.png)

Item       | Note
---------- | --------------------------------------------------------------------------------------------
Above Box  | Show information of precedent activities in this order: (ID) Relationship_Lead / Lag.
First Row  | Red means critical path and grey means non-critical path <br />Show activity ID and necessary duration
Second Row | Show the earliest start and the earliest finish
Third Row  | Show activity title
Forth Row  | Show the latest start and the latest finish

### Recommended Usage

In general, the process of planning is something like the following list. We recommend that you calculate total duration of your project by using this add-on in the fifth step.

1.  List all activities to complete your project
2.  Assign team members
3.  Calculate necessary duration for each activity
4.  Clarify relationship of activities by listing necessary input at the beginning and output at the end.&nbsp;
5.  Calculate total duration of your project by connecting activities based on the relationship.

### Restriction

* Do not change the name (list sheet and diagram sheet)
* Do not insert rows above the item row
* Do not edit or delete the hidden second row
* Do not insert columns between items

### License

GNU General Public License (GPL)

### Privacy Policy

we treat your privacy with respect and it is secured and will never be sold, shared or rented to third parties.

#### Information We Collect

In operating our add-on, we may collect and process the following data about you:

* Details of your visits to our website and the resources that you access, including, but not limited to, traffic data, location data, weblogs and other communication data
* Information that you provide by filling in forms on our website, such as when you registered for information or make a purchase
* Information provided to us when you communicate with us for any reason.
