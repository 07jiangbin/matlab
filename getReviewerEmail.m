     UA = {'Mozilla/5.0 (X11; U; Linux i686; en-GB; rv:1.8.1.6) Gecko/20070914 Firefox/2.0.0.7',
        'Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1667.0 Safari/537.36'};
    uaIdx = [1 2];
errPageNum = [];
S = [];
tic
for i = 1:1000
 
    try
        raw = urlread(['http://www.amazon.com/review/top-reviewers/ref=cm_cr_tr_link_' num2str(i) '?ie=UTF8&page=' num2str(i)],'useragent',UA{uaIdx(1)});
        TR = regexp(raw,'<tr id="reviewer.*?</tr>','match');
        %% 
        while numel(TR) == 0 
                   disp('切换UA尝试抓取');
                    uaIdx = uaIdx([2:end 1]);
                    raw = urlread(['http://www.amazon.com/review/top-reviewers/ref=cm_cr_tr_link_' num2str(i) '?ie=UTF8&page=' num2str(i)], 'useragent',UA{uaIdx(1)});
                     TR = regexp(raw,'<tr id="reviewer.*?</tr>','match');
        end
            
            
        
        %%
        for j = 1:numel(TR)
          tr = TR{j};
          rank = str2double(regexp(tr,'(?<=<td class="crNum"># )[\d|,]{1,}','match'));
          tmp.rank = rank;
          name = regexp(tr,'(?<=<b>).*?(?=</b>)','match');
          tmp.name = name{1};
          pUrl = regexp(tr,'(?<=<a href=").*?(?=")','match');
          tmp.pUrl = pUrl{1};
          crNum = regexp(tr,'(?<=<td class="crNum"> ).*?(?=</td>)','match');
          crNum = regexprep(crNum,'( |,)','');
          tmp.totalReviews = str2double(crNum{1});
          tmp.helpfulVotes = str2double(crNum{2});
          percentHelpful = regexp(tr,'(?<=<td class="crNumPercentHelpful"> )\d{1,}','match');
          tmp.percentHelpful = str2double(percentHelpful)/100;
          tmp.email = '';
          S = [S; tmp];
        end
        disp(['正在处理' num2str(i) '/' num2str(1000) '个排名页面,成功！共耗时' num2str(toc) '秒']);
    catch ex
        disp(['正在处理' num2str(i) '/' num2str(1000) '个排名页面,失败！共耗时' num2str(toc) '秒']);
        disp(['   出错原因：'  ex.message]);
        errPageNum = [errPageNum i];
    end 
end




