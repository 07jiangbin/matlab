%% ��������
Stores = {'freeshipping2you';'bestdailydeals';'orionmotortech';'seriouswholesaler';'globalfreeshipping';'wowfreeshipping1';'acautoparts04';'greaaatprice4u2';'green.wholesaler';'premiumautopart';'majorsavings14';'y-not69'};
conn = database('ERP', 'sa', 'Oceania0488'); 

curs = exec(conn,'select datediff(d,cast(MAX(DateTo) as DATE),GETDATE())-1 as date from oceaniaerp.dbo.DA_SKU50_Conversion');
curs = fetch(curs);
dat = curs.data;
N = dat.date;  % �ϴ����N�������

if N == 0 
    error('����������Ҫ����');
else 
    disp('��Ҫ�����������ڵ�����:');
    dateTo = (datenum(date)-N):(datenum(date)-1); % �������7�������
    disp(datestr(dateTo)); % չʾ��Ҫ���ص���������
end

% %% ����һ���ļ��б�
% dateFrom = datestr(datenum(date)-1,'yyyy-mm-dd');
%     F = dir(['*' recordDate '.xls']);
%     if numel(F)~=11
%         %���ȱʧ�ļ�
%         disp('δ�������µ�������:');
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
%     %% ��鷽��2
%          Stores2 = {};
%          for i = 1:numel(F)
%             tmp = regexp(F(i).name,'(?<=_).*?(?=_)','match');
%             Stores2 = [Stores2 tmp(1)];
%          end
%          disp(setdiff(Stores', Stores2)) 
%         error('����11������');
%     end

 %% ���������������
 
 sFile = [];
 lostFile ={};
 for i = 1 : numel(dateTo)
     for j = 1:numel(Stores)
        % tmpdateFrom =  datestr(dateTo(i)-6,'yyyy-mm-dd'); % ���ɰ��ܵ���ʼ����
         tmpdateFrom = datestr(dateTo(i),'yyyy-mm-dd');
         tmpdateTo =  datestr(dateTo(i),'yyyy-mm-dd');
         tmpFileName = ['Best Listing  ������_' Stores{j} '_' tmpdateFrom '_' tmpdateTo '.xls'];
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
     disp('ȱʧ�ļ���');
     disp(lostFile);
 else
     disp(['����' num2str(numel(sFile)) '���ļ�,�ļ���������!']);
 end
 
%%  Product, itemID, ConversionRate, SoldQty, Revenue, recordDate, IsOnline
curs = exec(conn, 'select SKU from OceaniaERP.dbo.DA_SKU50');
curs = fetch(curs);
data = curs.data;
skuList = data.SKU;
S = [];
for i = 1 : numel(sFile)
    disp(['���ڴ����' num2str(i) '/' num2str(numel(sFile)) '���ļ�...']);
    tmpS = parseXlsFile(sFile(i),skuList);  
    S = [S tmpS];
end
T = struct2table(S);
%disp('���!');
%% 
tableName = '[OceaniaERP].[dbo].[DA_SKU50_Conversion]';
colName = T.Properties.VariableNames;
fastinsert(conn, tableName,colName,T);
 exec(conn,'EXEC [OceaniaERP].[dbo].[DA_CVRReport_Daily]');


  