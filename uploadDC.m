%% 店铺名称
Stores = {'freeshipping2you';'bestdailydeals';'orionmotortech';'seriouswholesaler';'globalfreeshipping';'wowfreeshipping1';'acautoparts04';'greaaatprice4u2';'green.wholesaler';'premiumautopart';'majorsavings14';'y-not69'};
conn = database('ERP', 'sa', 'Oceania0488'); 

curs = exec(conn,'select datediff(d,cast(MAX(DateTo) as DATE),GETDATE())-1 as date from oceaniaerp.dbo.DA_SKU50_Conversion');
curs = fetch(curs);
dat = curs.data;
N = dat.date;  % 上传最近N天的数据

if N == 0 
    error('无新数据需要下载');
else 
    disp('需要下载以下日期的数据:');
    dateTo = (datenum(date)-N):(datenum(date)-1); % 生成最近7天的日期
    disp(datestr(dateTo)); % 展示需要下载的日期天数
end

% %% 最新一天文件列表
% dateFrom = datestr(datenum(date)-1,'yyyy-mm-dd');
%     F = dir(['*' recordDate '.xls']);
%     if numel(F)~=11
%         %检查缺失文件
%         disp('未下载如下店铺数据:');
% 
%     %%      for i = 1:numel(Stores)
%     %          tmp = [];
%     %        for j = 1:numel(F)
%     %            tmp = [tmp regexp(F(j).name,Stores{i},'match')];
%     %        end
%     %        if isempty(tmp)
%     %            disp(Stores{i});
%     %        end
%     %      end 
%     %% 检查方法2
%          Stores2 = {};
%          for i = 1:numel(F)
%             tmp = regexp(F(i).name,'(?<=_).*?(?=_)','match');
%             Stores2 = [Stores2 tmp(1)];
%          end
%          disp(setdiff(Stores', Stores2)) 
%         error('共有11个店铺');
%     end

 %% 处理最近七天数据
 
 sFile = [];
 lostFile ={};
 for i = 1 : numel(dateTo)
     for j = 1:numel(Stores)
        % tmpdateFrom =  datestr(dateTo(i)-6,'yyyy-mm-dd'); % 生成按周的起始日期
         tmpdateFrom = datestr(dateTo(i),'yyyy-mm-dd');
         tmpdateTo =  datestr(dateTo(i),'yyyy-mm-dd');
         tmpFileName = ['Best Listing  有流量_' Stores{j} '_' tmpdateFrom '_' tmpdateTo '.xls'];
         s = dir(tmpFileName);
         if numel(s) == 0 
             lostFile = [lostFile; tmpFileName];
         end
     tmpFile.dateFrom = tmpdateFrom;
     tmpFile.dateTo = tmpdateTo;
     tmpFile.fileName = tmpFileName;
     sFile = [sFile tmpFile];
     end
 end
 
 if numel(lostFile) > 0 
     disp('缺失文件！');
     disp(lostFile);
 else
     disp(['共有' num2str(numel(sFile)) '个文件,文件个数完整!']);
 end
 
%%  Product, itemID, ConversionRate, SoldQty, Revenue, recordDate, IsOnline
curs = exec(conn, 'select SKU from OceaniaERP.dbo.DA_SKU50');
curs = fetch(curs);
data = curs.data;
skuList = data.SKU;
S = [];
for i = 1 : numel(sFile)
    disp(['正在处理第' num2str(i) '/' num2str(numel(sFile)) '个文件...']);
    tmpS = parseXlsFile(sFile(i),skuList);  
    S = [S tmpS];
end
T = struct2table(S);
%disp('完成!');
%% 
tableName = '[OceaniaERP].[dbo].[DA_SKU50_Conversion]';
colName = T.Properties.VariableNames;
fastinsert(conn, tableName,colName,T);
 exec(conn,'EXEC [OceaniaERP].[dbo].[DA_CVRReport_Daily]');


  