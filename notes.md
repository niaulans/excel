## Notes

### Doing data preparation
```
- Collect, sort, and filter raw data
- Set the correct formatting and data types
- Clean raw data:
  - Remove duplicates
  - Correct errors
  - Fill in missing data
- Check data quality and summarize the data better

```

### Exploratory data analysis (EDA)
```
1. Data preparation 
  - Collecting
  - Cleaning
2. Data exploration 
  - Learn about each variable
  - Compute summary statistics
  - Find corelations and trends
  - Visualize the data
3. Hypothesis generation and further analysis
```

### Forecasting
```
- Process of predicting future outcomes and trends based on historical data usng statistical techniques and models
- Seasonality -> the correlation between the data and the time of year
- Bias -> the distortion of forecasting results from the way the analysis was set up
  - Sampling bias     -> data is collected in a way that is not representative
  - Confirmation bias -> only accepting results that the analyst already believes to be true
  - Anchoring bias    -> failing to adjust adequately for new data or changing trends
- Confidence interval -> the range within an actual outcome is likely to occur 
- Confidence level -> the probability an actual outcome is likely to fall within the confidence interval
- Weighted moving average -> assigns a weight to each data point based on its age
  - 3-month weighted moving average -> the average of the last three months with the most recent month weighted the most
- Simple moving average -> the average of a subset of data points
  - 3-month moving average -> the average of the last three months
```

### Choosing the Right Trendline
```
Depending on your data pattern:
1. Linear Trendline 📈
  If your data shows a constant increase or decrease.
  Example: Sales growing steadily every month.
  Formula: y=mx+b

2. Exponential Trendline 🚀
  If your data grows or declines exponentially (rapid increase or decrease).
  Example: Virus spread, app user growth.
  Formula: y = ae^(bx)
 
3. Polynomial Trendline 🔄
  If your data has fluctuations (ups and downs).
  Example: Stock prices, seasonal trends.
  Can use order 2, 3, or higher.
  Formula: y = ax^2 + bx + c (order 2)

4. Logarithmic Trendline 📊
  If your data rises quickly at first and then slows down.
  Example: Market growth that starts fast but stabilizes.
  Formula: y = a + b ln(x)

5. Moving Average Trendline 📉📈
  If your data is highly volatile and you want to see an average trend.
  Useful for financial or daily sales data.
```

### More complex charts
```
1. Single vs multi-series charts
  - Single series -> one data set
  - Multi-series -> multiple data sets
2. Combo charts
  - Combines two or more chart types
  - Useful for comparing different data types
3. Bullet chart
  - Compares a primary measure to one or more other measures
  - Useful for setting performance targets
  - Useful for comparing results vs benchmarks
4. Waterfall chart
  - Explains the net change in value between two data points
5. Scatter plot
  - Shows the relationship between two variables
  - Useful for identifying trends and correlations
```

### Data visualization best practices
```
Two dimensions or three dimensions?
- Two dimensions
  - Easy to interpret
  - Versatile and distraction-free

- Three dimensions
  - Hard to read the correct value
  - Perspective distorts the chart and confuses the eye

Labels, legends, and titles
- Not self-explanatory
- Axis does not start at 0 and misleading

Color: an ally or enemy?
- "Rainbow" charts - colors bring no value
- Avoid using red/amber/green in categorical legends
- Use color to draw attention to a data point

- Avoid redundant axes and chart titles
- Unnecessary legends
- Evaluate each chart element
- Bring focus to selected chart elements
- Use title and color to draw attention to the main message
- Use labels sparingly to highlight main data points
- Consider graying out or increase the transparency of less important data points
```

