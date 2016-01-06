function S = parseXlsFile(sFile, skuList)
if nargin == 1
    filterFlag = 0;
else
    filterFlag = 1;
end
dateFrom = sFile.dateFrom;
dateTo = sFile.dateTo;
fileName = sFile.fileName;
r = regexp(fileName,'(?<=有流量_).*?(?=_)','match');
Store = r{1};
S = []; 
        [~,~,raw] = xlsread(fileName);
        for j = 2:size(raw,1)
            product = raw{j,23};
            %
            if isnan(product)
                continue;
            end
            product = regexprep(product,'[.*','','ignorecase');
            product = regexprep(product,'-list.*','','ignorecase');
            SetOf = regexpi(raw{j,23},'(?<=\[listing\])\d{1,}(?=\[set\])|(?<=-listing-)\d{1,}','match');
            SetOf = str2double(SetOf);
            s.Product = product;
           s.Store = Store;
           s.ItemID = raw{j,29}(end-11:end);
           s.Title = raw{j,22};
           if ischar(raw{j,6})
               s.VisitorCnt = nan;
           else
               s.VisitorCnt = raw{j,6} ;  % 访客数
           end
           if ischar(raw{j,7})
               s.PageViewCnt = nan;
           else
               s.PageViewCnt =   raw{j,7}; % 页面浏览数
           end
           s.OrderCnt = raw{j,16};   % 订单数
           s.ConversionRate = str2double(raw{j,2}(1:end-1))/100; % 订单数/访客数
           s.SoldQty = raw{j,15}*SetOf;
           s.Revenue = str2double(raw{j,10}(2:end));   
           s.DateFrom = dateFrom;
           s.DateTo = dateTo;
           s.DateRange = datenum(dateTo) - datenum(dateFrom) + 1;
           s.IsOnline = raw{j,31} ;
           S = [S s];
        end
    if filterFlag %过滤出top50的sku
         S = skuFilter(S, skuList);     
    end
end

function S2 = skuFilter(S, skuList)
S2 = [];
   for i = 1:numel(S)
        for j = 1:numel(skuList)
            if isequal(S(i).Product,skuList{j})
                S2 = [S2 S(i)];
                break;
            end
        end
   end
end