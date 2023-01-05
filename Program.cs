// See https://aka.ms/new-console-template for more information
using MiniExcelLibs;
using MiniExcelLibs.Attributes;

Console.WriteLine("Hello, World!");
var path= Directory.GetCurrentDirectory();
var input=System.IO.Path.Combine(path,"input","12月工时.xlsx");
var output=System.IO.Path.Combine(path,"output","output-12月工时-1.xlsx");
var rows = MiniExcel.Query<InputItem>(input).ToList();
var rd = new Random();
var month=12;
Console.WriteLine(rows.Count);
var ouputList = new List<OutputItem>(); //
foreach(var row in rows){
    for(var i=1;i<=31;i++){
        var fieldName=$"D{i}";
        var fieldValue = row.GetType().GetProperty(fieldName).GetValue(row,null);
        if(fieldValue!=null){
            var hours = Convert.ToDouble(fieldValue);
            var minutes = rd.Next(20,29);
            var seconds = rd.Next(0,59);
            var dt1= new DateTime(2022,month,i,8,minutes,seconds);

            var seconds1 = rd.Next(0,40);
            var minutes1 = rd.Next(31,36);
            var dt2=new DateTime(2022,month,i,17,minutes1,seconds1);

           if(hours==10.5){
                seconds1 = rd.Next(0,60);
                 minutes1 = rd.Next(31,45);
                 dt2=new DateTime(2022,month,i,20,minutes1,seconds1);
            }
            else if(hours==8) {
                Console.WriteLine(dt2);
            }
            else if(hours==7.5){
                 seconds1 = rd.Next(0,60);
                 minutes1 = rd.Next(0,5);
                 dt2=new DateTime(2022,month,i,17,minutes1,seconds1);
            }
            else if(hours>8 ){
                hours=hours+2;
                 seconds1 = rd.Next(0,30);
                 minutes1 = rd.Next(0,2);
                 dt2=dt1.AddHours(hours).AddMinutes(minutes1).AddSeconds(seconds1);
            }
            else if(hours<8 && hours>4){
                hours=hours + 1.5;
                seconds1 = rd.Next(0,30);
                minutes1 = rd.Next(0,2);
                dt2=dt1.AddHours(hours).AddMinutes(minutes1).AddSeconds(seconds1);
            }else {
                seconds1 = rd.Next(0,30);
                minutes1 = rd.Next(0,2);
                dt2=dt1.AddHours(hours).AddMinutes(minutes1).AddSeconds(seconds1);
            }
            
            
            
            var output1=new OutputItem(){
                Name=row.Name,
                WorkNo=row.WorkNo,
                Supply=row.Supply,
                SupplyId=row.SupplyId,
                Department=row.Department,
                Job=row.Job,
                Date1 = dt1.ToString("yyyy-MM-dd HH:mm:ss"),
            };
            var output2=new OutputItem(){
                Name=row.Name,
                WorkNo=row.WorkNo,
                Supply=row.Supply,
                SupplyId=row.SupplyId,
                Department=row.Department,
                Job=row.Job,
                Date1 = dt2.ToString("yyyy-MM-dd HH:mm:ss"),
            };

            ouputList.Add(output1);
            ouputList.Add(output2);
        }
    }
}

Console.WriteLine(ouputList.Count);
MiniExcel.SaveAs(output,ouputList);
public class InputItem{
    [ExcelColumnName("序号")]
    public string? Id{get;set;}
    [ExcelColumnName("工号")]
    public string? WorkNo{get;set;}
    [ExcelColumnName("姓名")]
    public string? Name{get;set;}
    [ExcelColumnName("供方")]
    public string? Supply{get;set;}
    [ExcelColumnName("供方ID")]
    public string? SupplyId{get;set;}
    [ExcelColumnName("部门")]
    public string? Department{get;set;}
    [ExcelColumnName("岗位")]
    public string? Job{get;set;}
    [ExcelColumnName("1")]
    public decimal? D1{get;set;}
    [ExcelColumnName("2")]
    public decimal? D2{get;set;}
    [ExcelColumnName("3")]
    public decimal? D3{get;set;}
    [ExcelColumnName("1")]
    public decimal? D4{get;set;}
    [ExcelColumnName("5")]
    public decimal? D5{get;set;}
    [ExcelColumnName("6")]
    public decimal? D6{get;set;}
    [ExcelColumnName("7")]
    public decimal? D7{get;set;}
    [ExcelColumnName("8")]
    public decimal? D8{get;set;}
    [ExcelColumnName("9")]
    public decimal? D9{get;set;}
    [ExcelColumnName("10")]
    public decimal? D10{get;set;}
    [ExcelColumnName("11")]
    public decimal? D11{get;set;}
    [ExcelColumnName("12")]
    public decimal? D12{get;set;}
    [ExcelColumnName("13")]
    public decimal? D13{get;set;}
    [ExcelColumnName("14")]
    public decimal? D14{get;set;}
    [ExcelColumnName("15")]
    public decimal? D15{get;set;}
    [ExcelColumnName("16")]
    public decimal? D16{get;set;}
    [ExcelColumnName("17")]
    public decimal? D17{get;set;}
    [ExcelColumnName("18")]
    public decimal? D18{get;set;}
    [ExcelColumnName("19")]
    public decimal? D19{get;set;}
    [ExcelColumnName("20")]
    public decimal? D20{get;set;}
    [ExcelColumnName("21")]
    public decimal? D21{get;set;}
    [ExcelColumnName("22")]
    public decimal? D22{get;set;}
    [ExcelColumnName("23")]
    public decimal? D23{get;set;}
    [ExcelColumnName("24")]
    public decimal? D24{get;set;}
    [ExcelColumnName("25")]
    public decimal? D25{get;set;}
    [ExcelColumnName("26")]
    public decimal? D26{get;set;}
    [ExcelColumnName("27")]
    public decimal? D27{get;set;}
    [ExcelColumnName("28")]
    public decimal? D28{get;set;}
    [ExcelColumnName("29")]
    public decimal? D29{get;set;}
    [ExcelColumnName("30")]
    public decimal? D30{get;set;}
    [ExcelColumnName("31")]
    public decimal? D31{get;set;}
    [ExcelColumnName("工时合计")]
    public decimal? Total{get;set;}
   
}

public class OutputItem{
public string? Id{get;set;}
    [ExcelColumnName("工号")]
    public string? WorkNo{get;set;}
    [ExcelColumnName("姓名")]
    public string? Name{get;set;}
    [ExcelColumnName("供方")]
    public string? Supply{get;set;}
    [ExcelColumnName("供方ID")]
    public string? SupplyId{get;set;}
    [ExcelColumnName("部门")]
    public string? Department{get;set;}
    [ExcelColumnName("岗位")]
    public string? Job{get;set;}
    [ExcelColumnName("打卡时间")]
    public string? Date1{get;set;}
}