clc,clear
close all
filename = '第9周';
classcount = '课程统计用.xlsx';
xlsname = '考勤报表.xls';
subxlsname = '考勤汇总表';
writename = '结果.xls';
% untar(strcat(rarname,'.rar'),rarname)
%%
[~,~,class_content]=xlsread(classcount);
[len,~] = size(class_content);
numlist=zeros(len-1,1);
stu_name = cell(len-1,1);
class = cell(len-1,1);
for i = 2: len
    numlist(i-1) = i-1;
    stu_name{i-1} =class_content{i,1};
    class{i-1} = class_content{i,2};
end
%%
result={};
xlsname_path = strcat(filename, '\',xlsname)
[~,~,xls_content]=xlsread(xlsname_path,subxlsname);
[len,~] = size(xls_content);

if len>length(class)
    class{length(class)+1}=NaN;
end
num=1;
for i = 5: len
    students_num = str2double(xls_content{i,1});
    students_name = xls_content{i,2}
    group = xls_content{i,3};
    workhours = str2double(xls_content{i,5});
    addhours = str2double(xls_content{i,10});
    sumhours = workhours + addhours;
    result{num,1} = students_num;
    result{num,2} = students_name;
    result{num,3} = sumhours;
    result{num,4} = class{num};
    num=num+1;
end

xlswrite(strcat(filename,'\',writename),result)