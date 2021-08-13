# outlook_report_automation
Generates reports from attached csv files

## Goal of this project
To automate report generation from attached csv files. 

## How it works
By receiving an Email a subject of the Email will be checked. If it contains key phrase (in our case it is: "Grafana Reporting") -> extract/transform of attached
csv file will be executed to perform report generation. 

## Data Structure of attached csv file
Attached csv file represents Time Series Data, which dowloaded from Grafana. It has following columns:
1. **Time** - date time in either (%Y-%m-%d %H:%M:%S) or in unix (ms) format.
2. **Measured Data** - rest of the columns (may be several) represent measured data
### Sample data of csv file
|Time               |System    |User      |Iowait |Total |
| ----------------- | -------- | -------- | ----- | ---- |
|2021-08-11 10:02:30| 0.0997   | 0.0550   | 0.056 | 0.159|

## Report generation
For each mesured data column a worksheet in excel workbook will be created with plotted graph and summarized values(in this case: max/min/avg) and saved to output directory.

### Example of report

![Alt text](./Report.png?raw=true "Title")

## References
Some of the code(outlook listener) was taken from [this repo](https://gist.github.com/burdenless/fd2c92e468a3d07f5c37) 


