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


# Chapter 3: Switch Row/Column Data

## Switch Row/Column Data

When we entered our data into our worksheet, we arranged it in a columnar fashion with the months down one column and their corresponding sales numbers in the adjacent column. Excel interpreted the data this way and created the charts accordingly. But sometimes Excel gets it wrong and assumes you assembled the data in a row fashion. If you create your chart and the information on the X and Y axes appears to be switched, here's a simple solution that will correct the problem.

Before you begin, let's add a legend to the chart because it will help illustrate the switch row/column concept when you see what each colored column represents before and after the change. Make sure your chart is active, and select the Chart Elements option (plus sign icon) to the right of your chart. Place a check mark next to **Legend**. A legend appears to the right of the plot area in your chart. As you can see, the blue columns represent sales. Also note that the horizontal axis is a list of the various months in your data. Click the **Chart Elements** icon again to hide the Chart Elements menu options.

From the Chart Tools Ribbon, on the **Design** tab, locate the section labeled **Data**. (Remember, if the Chart Tools Ribbon is missing, simply click anywhere on the chart to bring it back into view.)

Now, from the Data section, click **Switch Row/Column** to switch the orientation of the two axes. As you can see, the colored columns represent the months (look at your legend) and the horizontal axis row represents sales. The height of each column still represents the sales value, so they're both correct. One just gives you a different look at the data. In the example, the changes are pretty much just cosmetic. But as you experiment with this option on your own data, you'll encounter changes that may be dramatic or may produce a result that doesn't tell the story you want it to. So, please be cautious.

Click the Switch Row/Column icon again, and your chart will snap back to its original state. This icon acts as a toggle, so if you ever get a chart that doesn't tell the story you were hoping, give the Switch Row/Column icon a try.

## Titles
Adding a title to your chart is as simple as the steps to add a legend. Click the Chart Elements option (plus sign), and place a check mark next to **Chart Title**. Click the **Chart Elements** icon again to close the menu.

### Editing Chart Text
The boss will probably want a title other than the default Chart Title, so let's change that. Notice that the title has appeared at the top of the chart area, and there are lines and circles in each of the corners. This means the title is currently active. If you don't have this border anymore, it probably means you clicked elsewhere on your worksheet tab and the title is no longer the active item. To remedy this, simply click the chart title to activate it again. With the title active, just begin typing a new title. It's that simple. As you're typing, you won't see the chart title follow along, but if you look in the formula bar, you'll see what you're typing. Go ahead and type the words Monthly Sales as your new title. When done, press ENTER.

Here's how it should look after the chart title update:



## Data Labels
Data labels can add a little flair to your chart plot area. Data labels are normally numbers, but you can also display them as percentages or even as category names. It all depends on the type of chart you're working with. Each of the different chart types may have different options. As I mentioned earlier, I chose a chart type with data labels because that's my preference. But the default font color or format may be less than ideal. Here's how to edit them:

Select any one of the numeric data labels. By default, all of them are selected and now have a border around them. Go to the Home tab and select a new font color, like black or something dark. Instantly, your data labels within your chart reflect the new font color.

Be aware that you can also change the font color and other elements of your chart by going to the Format tab and utilizing options in the Shape Styles and WordArt Styles groups. Since the exercise above was a simple font color change, by habit I always do these from the **Home** tab since all of my nonchart formatting changes are done here as well. But feel free to use the option that works best for you.

# Chapter 4: Chart Formatting
## Chart Formatting
The number of options to format a chart is nearly limitless, and to add to the confusion, there are different ways to get to these options. For example, you've already completed exercises where you've used the context-specific chart tabs in the Ribbon, the Chart Elements icon next to the chart, as well as the font options on the Home tab.

As I mentioned earlier, you'll find the more complex or obscure options in the Ribbon. In fact, all chart options will be in the Ribbon somewhere. I tend to go to the three icons to the right of my chart first. If the option I need isn't there, then I go to the Ribbon.

## Current Selection
On the Format tab, the top option in the Current Selection area is Chart Elements. Click the drop-down arrow to see all the different chart elements available for your chart type. Choose an item, such as Chart Title. Notice the title in your chart is now selected and has a border around it. By choosing the appropriate option in the Chart Elements list, you can quickly select the corresponding item in the chart itself.

The reverse also works. If you click an element on the chart, it will change the option in the Chart Elements list to match. To see how this works, use your cursor to select your chart legend. When you click the legend, you'll see the Chart Elements list changes to match the item you selected.

I point this out because the Chart Elements list can help you identify any component of your chart. It can also provide you with a visual cue in case you click a piece of your chart and wonder if you've selected the right item. When you click the item on the chart, simply look up into the Chart Elements list and see what's displayed.




# Q&A
Q: The Chart Tools Ribbon disappears from time to time. How do I make it visible again?

A: Simple! Just click any chart or graph in your workbook, and Excel will automatically retrieve the Chart Tools Ribbon for you to use.

Q: Why don't you insert a space between words on your sheet names? Wouldn't that make them easier to read?

A: Actually, Excel will allow you to use spaces in sheet names. However, it can make your life a little more difficult if you later decide that you want to write a formula using values scattered across two or more different sheets. Keeping spaces out of the sheet names will make it much easier to write those kinds of formulas, should the need arise.
