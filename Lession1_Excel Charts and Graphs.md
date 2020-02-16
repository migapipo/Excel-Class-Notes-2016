# Chapter 2: Creating a Chart

## Charting Options Using Excel's Ribbon
In Excel, there are many charting options available on the Ribbon as well. We could spend all day going through them all, so I'll just touch on the ones I think will be the most useful to you. When you get some free time, be sure to experiment with the others. If you run into trouble or want further explanation on a certain topic, just leave me a post in the Discussion Area. I'll be happy to help you!

With your line chart selected, click the **Design** tab. Notice the labels underneath each grouping of icons: Chart Layouts, Chart Styles, Data, Type, and Location.


**Chart Layouts** : Here you can add or modify any chart element that pertains to your chart type by selecting the Add Chart Element option. Or select the Quick Layout option to apply a predefined chart layout that includes or suppresses various combinations of chart elements.

**Chart Styles**: Here you can modify the color and style of your chart. Click the bottom down arrow at the right side of the Chart Styles section to see all available style options for line charts. There are 16 different chart styles to pick from. The current style is highlighted with a green frame.

**Chart Types**: Luckily, the developers at Microsoft realized that no matter how smart Excel users are, they're bound to change their minds! Now you don't have to go through all the steps again to create a new chart, saving you tons of time. Microsoft Excel will quickly and easily do the hard work for you. Let's say you included this line chart in a report you sent to the boss. Although everyone agrees that sales appear to be dropping, they just don't like line charts, so your boss has requested a different chart type to tell the story. No reason to worry. It's simple to change a chart to a new type.


## Changing the Chart Type
Located to the right of the *Chart Styles* and *Data* group, you'll notice the Type group. Within that group is just one icon labeled *Change Chart Type*. Click the icon labeled **Change Chart Type**. The Change Chart Type dialog box appears.


[pic1]

By default, the dialog box opens to the All Charts tab. By selecting different options from the list on the left, different chart previews will display in the main area. If a chart doesn't have a preview, that's Excel's way of telling you the data doesn't work with that chart type.

Next, click the Recommended Charts tab. Look familiar? The four original charts are listed on the left. By default, Excel chose the Clustered column chart type.

This seems like a great substitute to the original because it conveys the right message, and the boss will be happy because it's not a line chart. With that chart selected, click OK.

Returning to the work area, you'll see that you've easily changed the Line chart to a Clustered column chart that looks like this:

[pic2]


Wow—look at that drop in sales! This chart also shows that sales have been declining since January, and that May was a terrible month. I think the boss will be happy with the alternative chart provided, but she probably won't be happy about the sales numbers.


## Chart Updates
What would happen to the chart if you need to change the chart input area? For example, let's say that the accounting department provided the wrong sales number for the month of May—the true number should be $25,000. Use the scroll bar to scroll back up to the top of your worksheet where you input your data. Change cell B6 from $10,905 to $25,000, and watch your chart automatically update with the new number. Excel keeps that link between the data area and the chart indefinitely. You just have to make sure the numbers and labels are correct, and Excel will do the rest.

Now that you've seen how Excel automatically updates your chart when you change a number, please change the amount in cell B6 back to 10,905 before moving on.



## Move Chart Location
By default, when Excel created your chart, it embedded it as an object in the worksheet. This may or may not be what you want. For example, let's say your boss wants the chart to appear on a separate workbook tab. You could create an additional workbook tab and then copy and paste the chart to the new location, but there's an easier way.

From the Chart Tools Ribbon, click the **Design** tab, and then select **Move Chart**. The Move Chart dialog box appears. The option currently selected should be Object in. On the right, you'll see Sheet1. If you want to move the chart to another preexisting sheet tab, select the drop-down arrow, choose the appropriate sheet tab, and then click **OK**.

Or you can select the New Sheet option. The field to the right is an input field. You can either use the sheet tab name already provided or type your own. Click OK, and Excel will create a new tab with just the chart on it. Please be aware, when a chart has been moved to its own chart sheet, the Enable Live Preview and some editing options are no longer available. For example, when a chart is on its own sheet, you won't be able to change the size of the chart.

[Move Chart dialog box]

If you move your chart as an object to another sheet or as a new sheet and don't like the result (say, your boss changed her mind), use the Move Chart Location option to change it back.
