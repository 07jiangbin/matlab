S = {'acautoparts04','http://www.datacaciques.com/store/report/3351210760';
    'bestdailydeals','http://www.datacaciques.com/store/report/1694108836';
    'freeshipping2you','http://www.datacaciques.com/store/report/1313909919';
    'globalfreeshipping','http://www.datacaciques.com/store/report/2858303830';
    'greaaatprice4u2','http://www.datacaciques.com/store/report/3449860238';
    'green.wholesaler','http://www.datacaciques.com/store/report/3480403037';
    'majorsavings14','http://www.datacaciques.com/store/report/4221355050';
    'orionmotortech','http://www.datacaciques.com/store/report/1772951905';
    'premiumautopart','http://www.datacaciques.com/store/report/3608735570';
    'seriouswholesaler','http://www.datacaciques.com/store/report/2494045869';
    'wowfreeshipping1','http://www.datacaciques.com/store/report/2924723610';
    'y-not69','http://www.datacaciques.com/store/report/3380993560'};

ie = actxserver('internetexplorer.application');
ie.visible = 1;
ie.Navigate(S{1,2});

% ѡ��������
while ~strcmp(ie.ReadyState,  'READYSTATE_COMPLETE')
    pause(0.5);
end
 %t = ie.document.body.getElementsByClassName('yesterday');

 t = ie.document.body.getElementsByClassName('today');

t.item(0).click

%%

ws = actxserver('WScript.Shell');
%{
ws.SendKeys('%{TAB}')  % ���� alt+tab��
ws.SendKeys('^S'); % ���� CTRL+ S

%}
for i = 1:size(S,1)
    ie.Navigate(S{i,2});
    pause(5)
    while ~strcmp(ie.ReadyState,  'READYSTATE_COMPLETE')
    pause(2);
    end
    d = ie.document.body.getElementsByClassName('download-btn btn');
    d.item(1).click 
    while ~strcmp(ie.ReadyState,  'READYSTATE_INTERACTIVE')
    pause(2);
    end  
%     pause(5)
%    % ws.SendKeys('%{TAB}')  % ���� alt+tab��
%     s = ws.AppActivate([ie.LocationName ' - Internet Explorer']); 
%     pause(0.5);
%      ws.SendKeys('%S'); % ���� alt + S
%    while ~s
%      s = ws.AppActivate([ie.LocationName ' - Internet Explorer']);
%      pause(0.5);
%      ws.SendKeys('%S'); % ���� alt + S
%      pause(0.5);
%    end
disp(['�����ص�' num2str(i) '���ļ�,�˺�:' S{i,1} '�������������...']);
pause
end











