% 导入athletes.csv
athletes = readtable("summerOly_athletes.csv");
% 导入host.xlsx
% 设置导入选项并导入数据
opts = spreadsheetImportOptions("NumVariables", 3);

% 指定工作表和范围
opts.Sheet = "summerOly_hosts";
opts.DataRange = "A2:C36";

% 指定列名称和类型
opts.VariableNames = ["Year", "HostCountry", "HostCity"];
opts.VariableTypes = ["double", "string", "string"];

% 指定变量属性
opts = setvaropts(opts, "HostCountry", "WhitespaceRule", "preserve");
opts = setvaropts(opts, ["HostCountry", "HostCity"], "EmptyFieldRule", "auto");

% 导入数据
host = readtable("D:\MATLAB_files\2025spring\2025MCM&ICM\2025MCM正赛\2025_Problem_C_Data\host.xlsx", opts, "UseExcel", true);

% 清除临时变量
clear opts
% 导入medal_counts.csv
medal_counts = readtable("medal_counts.csv");
% 导入 program.xlsx
% 设置导入选项并导入数据
opts = spreadsheetImportOptions("NumVariables", 35);

% 指定工作表和范围
opts.Sheet = "summerOly_programs";
opts.DataRange = "A2:AI77";

% 指定列名称和类型
opts.VariableNames = ["Sport", "Discipline", "Code", "SportsGoverningBody", "year_1896", "year_1900", "year_1904", "year_1906", "year_1908", "year_1912", "year_1920", "year_1924", "year_1928", "year_1932", "year_1936", "year_1948", "year_1952", "year_1956", "year_1960", "year_1964", "year_1968", "year_1972", "year_1976", "year_1980", "year_1984", "year_1988", "year_1992", "year_1996", "year_2000", "year_2004", "year_2008", "year_2012", "year_2016", "year_2020", "year_2024"];
opts.VariableTypes = ["string", "string", "string", "categorical", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];

% 指定变量属性
opts = setvaropts(opts, ["Sport", "Discipline", "Code"], "WhitespaceRule", "preserve");
opts = setvaropts(opts, ["Sport", "Discipline", "Code", "SportsGoverningBody"], "EmptyFieldRule", "auto");

% 导入数据
program = readtable("D:\MATLAB_files\2025spring\2025MCM&ICM\2025MCM正赛\2025_Problem_C_Data\program.xlsx", opts, "UseExcel", false);

% 清除临时变量
clear opts

noc = string(athletes.NOC);
country = string(athletes.Team);
NOC2Country = ["",""];
k = 0;
tic
for n = 1:length(noc)
    if any(any(contains(NOC2Country,noc(n))))
        continue
    end
    temp = [noc(n),country(n)];
    if any(any(contains(temp(2),'/'))) || any(any(contains(temp(2),'-')))
        continue
    end
    k = k + 1;
    NOC2Country(k,:) = temp;

end
clear k n temp


