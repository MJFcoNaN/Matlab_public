% Find correspondence between some excel sheets...
% 
% Created by: MJFcoNaN
% E-mail    : mjfconan@outlook.com
% 20-Mar-2017 20:43:02
%
% Every excel sheets must have these columns 
% and the 1st row must at least have these certain names.
% 
% There could be more unused columns, and the order of all columns will be 
% processed correctly automatically.
% 
% I am lazy to separate the only defined function of
% "deal_with_ssc_xls_file", which makes this script cannot run sector by
% sector.
% 
% =========================================================================
% kongbai.xlsx: ("ttt2" is not considered now)
%
%  only one sheet, named "all"
%
%               filter_id	ttt         ttt1        blank   ttt2
%               7600        2017/1/16	14:22:29	0.7618 	g
%               7601        2017/1/16	14:23:32	0.7839 	g
%               7602        2017/1/16	14:23:50	0.7755 	g
%               7603        2017/1/16	14:24:05	0.7696 	g
%               7604        2017/1/16	14:24:20	0.7757 	g
%               ...
% =========================================================================
% zongzhong.xlsx: the same as kongbai.xlsx
%
% only one sheet, named "all"
% 
%               filter_id	ttt         ttt1        add     ttt2
%               6200        2017/1/18	9:35:55     0.7849	g
%               6201        2017/1/18	9:36:28     0.796	g
%               6202        2017/1/18	9:37:39     0.802	g
%               6203        2017/1/18	9:37:56     0.7985	g
%               6204        2017/1/18	9:38:04     0.8357	g
%               6205        2017/1/18	9:38:13     0.862	g
%               6206        2017/1/18	9:38:22     0.803	g
% 
% =========================================================================
% 盐度、悬浮物测定记录表2.xlsx
% 
% ##### #####  CAUSTIONS 1 #####  #####  
% "sample_id" is the number sticked on the glass bottles, 
% and may be duplicate, for example, many ships.
% Therefore, they should be separated into many sheets, and the sheets'
% name will be used to recognize them by add a certain value:
% "color_list", "color_value"
%
% ##### #####  CAUSTIONS 2 #####  ##### 
% If there is more than one filter for any "sample_id", 
% you must delete others and only left one. ("merge to one filter")
% The same step should be done in "kongbai.xlsx", "kongbai.xlsx"
% for example, in "kongbai.xlsx", "1111" is the sum weight for 3 filters,
% because these 3 filters corresponde to one water sample.
% 
%              filter_id	ttt         ttt1        add     ttt2
%              1111        2017/1/16	20:20:20	2.22 	g
%
% similar in "zongzhong.xlsx",
% 
%              filter_id	ttt         ttt1        add     ttt2
%              1111        2017/1/26	10:10:10	11.11 	g
%
% ##### #####          #####  #####  
%
%   sheet "light_blue"
%
%               sample_id	salt        filter_id
%               281         0.29 		7592
%               282         0.39 		7593
%               429         0.36 		7796
%               351         1.47 		5442
%               294         1.02 		8177
%               306         0.40 		8189
%               396         0.62 		6087
%               1           3.13 		7948
%               2           3.25 		7949
%
%   sheet "green"
%
%               sample_id	salt	水样容积	filter_id	滤纸号2	滤纸号3
%               1           0.1884              8872					
%               2           0.1884              8873					
%               3           0.1867              8874			
% =========================================================================
% "汇总.xlsx"
%  Only one sheet, namely "all", will be considered.
%  Missing data is OK, 
%  then the sample_id is the base value such as 10000, 20000.
%  only need {'site','tide','date', 'hour', 'h','layer','sample_id'},
%  and other columns will be discarded.
%  ##### #####  CAUSTIONS  #####  #####  
%  "sample_id" must be the final id, 
%  means, added by base value assigned in the pair of 
%  "color_list", "color_value"
% ##### #####       #####  #####
%
%           site       tide     date        hour	h	layer 瓶号 sample_id
%           南槽上       小     2016/12/6       8	10     0    2   20002 
%           南槽上       小     2016/12/6       8	10     0.2  3   20003 
%           南槽上       小     2016/12/6       8	10     0.4  4   20004 
%           B3          大     2016/12/16     17   5.6    0.6      10000 
%           B3          大     2016/12/16     17	  5.6    0.8      10000 
%           B3          大     2016/12/16     17	  5.6    1    375 10375 
%           B3          大     2016/12/16     18	  4.2    0    376 10376 
%           北港中       小     2016/12/8       9	8.6    0.6  184 30184 
%           北港中       小     2016/12/8       9	8.6    0.8  185 30185 
%           北港中       小     2016/12/8       9	8.6    1    186 30186 
%
% =========================================================================
%% ========================================================================

clear,clc

sample_volume = 0.6; % 600 mL
base_value    = 1E4;
color_list    = {'red', 'white', 'dark_blue', 'light_blue'};
color_value   = [  1,     2,        3,            4]  * base_value;

data_dir      = '/Users/mjfconan/SSC水样';
% blank
xls_ssc_00    = fullfile(data_dir, 'kongbai.xlsx');
% added filters
xls_ssc_01    = fullfile(data_dir, 'zongzhong.xlsx');
xls_salinity  = fullfile(data_dir, '盐度、悬浮物测定记录表2.xlsx');
xls_position  = fullfile(data_dir, '汇总.xlsx');

% if weight again and "zongzhong" has an added sheet named "new_add"
is_re_weitht              = false; 
% if you screw up in certain days, you may delete them by time...
is_delete_some_days_value = false;
xls_out = fullfile(data_dir, 'outout.xlsx');

%% ========================================================================
%% original xlsx data from balance
tbl_ssc_00 = deal_with_ssc_xls_file(xls_ssc_00, 'all'); % blank
tbl_ssc_01 = deal_with_ssc_xls_file(xls_ssc_01, 'all'); % add sediment
if is_re_weitht
    tbl_ssc_02 = deal_with_ssc_xls_file(xls_ssc_01, 'new_add'); % new add
else
    % the same as tbl_ssc_01, but change variable "add" to "new_add"
    tbl_ssc_02 = tbl_ssc_01;
    tbl_ssc_02.Properties.VariableNames{'add'} = 'new_add';
end

% delete data == nan
tbl_ssc_00(isnan(tbl_ssc_00.blank  ), :) = [];
tbl_ssc_01(isnan(tbl_ssc_01.add    ), :) = [];
tbl_ssc_02(isnan(tbl_ssc_02.new_add), :) = [];

%%
% valid data
tbl_ssc_new = innerjoin(tbl_ssc_00, tbl_ssc_02, 'Keys', 'filter_id');
tbl_ssc_new.ssc = (tbl_ssc_new.new_add - tbl_ssc_new.blank)/sample_volume;
tbl_ssc_new.Properties.VariableNames{'new_add'} = 'add';
% tbl_ssc_new.datetime(:) = datetime(2017,6,10);
tbl_ssc_new = ...
    tbl_ssc_new(:, {'filter_id', 'blank', 'add', 'datetime', 'ssc'});

tbl_ssc_11      = innerjoin( tbl_ssc_00, tbl_ssc_01, 'Keys', 'filter_id');
tbl_ssc_11.diff = tbl_ssc_11.add - tbl_ssc_11.blank;
tbl_ssc_11.ssc  = tbl_ssc_11.diff / sample_volume;
tbl_ssc_11      = tbl_ssc_11(:, {'filter_id', 'blank', 'add',...
    'datetime_tbl_ssc_01', 'ssc'});
% index of awful data
if is_delete_some_days_value
    ind_awful = ...
        ( tbl_ssc_11.datetime_tbl_ssc_01.Month ==   2    ...
        & tbl_ssc_11.datetime_tbl_ssc_01.Day   ==  14) | ...
        ( tbl_ssc_11.datetime_tbl_ssc_01.Month ==   3    ...
        & tbl_ssc_11.datetime_tbl_ssc_01.Day   ==  14);
    % delete awful data
    tbl_ssc_11(ind_awful, :) = [];
end
% rename datetime
tbl_ssc_11.Properties.VariableNames{'datetime_tbl_ssc_01'} = 'datetime';
%% only add new data where is blank
[lia, locb] = ismember(tbl_ssc_new.filter_id, tbl_ssc_11.filter_id);
% only add
tbl_ssc_11 = [tbl_ssc_11; tbl_ssc_new(~lia, :)];
% sort
tbl_ssc_11 = sortrows(tbl_ssc_11, 'filter_id');
%
tbl_ssc = tbl_ssc_11;
% positon
tbl_pos = readtable(xls_position, 'Sheet', 'all');
tbl_pos = ...
    tbl_pos(:, {'site','tide','date', 'hour', 'h','layer','sample_id'});
% sample id == "zero", NO data
tbl_pos(mod(tbl_pos.sample_id, base_value)==0, :)=[];
%% sample
tbl_sample = table;
for ii = 1:length(color_list)
    tmp_color = color_list{ii};
    tmp_color_value = color_value(ii);
    tbl_salt = readtable(xls_salinity, 'Sheet', tmp_color);
    tbl_salt = tbl_salt(:, {'sample_id', 'filter_id', 'salt'});
    tbl_salt.sample_id = tbl_salt.sample_id + tmp_color_value;
    tbl_sample = [tbl_sample; tbl_salt];
end
clear tmp_*
tbl_sample(isnan(tbl_sample.filter_id), :)=[];


%%
tbl_1        = innerjoin(      tbl_sample, tbl_ssc, 'Keys', 'filter_id');
tbl          = innerjoin( tbl_pos   , tbl_1  , 'Keys', 'sample_id');
tbl.datetime = ...
    datetime(datenum(tbl.date)+tbl.hour/24, 'ConvertFrom','datenum');
%%
for rl = 0:0.2:1
    s_cor_std = sprintf('%+2.1f', rl);
    s_cor_std = s_cor_std([2,4]);
    ind       = (abs(tbl.layer - rl)<1e-2);
    
    tbl.(['salt_', s_cor_std])      = nan(length(tbl.sample_id),1);
    tbl.(['ssc_' , s_cor_std])      = nan(length(tbl.sample_id),1);
    
    tbl.(['salt_', s_cor_std])(ind) = tbl.salt(ind);
    tbl.(['ssc_' , s_cor_std])(ind) = tbl.ssc(ind);
end
% sort table
tbl       = sortrows(tbl, {'site', 'datetime'});

tmp_tbl   = tbl(:,{'site','datetime'});
[C,ia,ic] = unique(tmp_tbl,'rows');
%%
for ii = 1:length(ia(1:end-1))
    ind_first = ia(ii);
    ind_end   = ia(ii+1)-1;
    tbl_line  = tbl(ind_first:ind_end, :);
    % "average"
    tmp_data          = tbl_line{:, end-11:end};
    tmp_data(1,:)     = mean(tmp_data, 1, 'omitnan');
    tmp_data(2:end,:) = nan;
    % only preserve the 1st row, other rows will be set as NaN
    tbl_line{:, end-11:end} = tmp_data;
    tbl(ind_first:ind_end, :) = tbl_line;
    
end

tbl = tbl(ia,:);
% delete unused, replicate or misguiding variables
tbl(:, {'date', 'hour', 'layer', 'salt', 'ssc'}) = [];

%% write out
writetable(tbl, xls_out)

% functions
function tbl_out = deal_with_ssc_xls_file(xls_path, xls_sheet_name)
tbl_out = readtable(xls_path, 'Sheet', xls_sheet_name); % blank

tbl_out.datetime = tbl_out.ttt + tbl_out.ttt1;

tbl_out.ttt = []; tbl_out.ttt1 = []; tbl_out.ttt2 = [];

tbl_out.datetime.Format = 'yyyy-MM-dd HH:mm:ss';
% unify
tbl_out = unique(tbl_out, 'rows');

% sort
tbl_out = sortrows(tbl_out, 'filter_id');

% delete filterid == nan
tbl_out(isnan(tbl_out.filter_id),:)=[];
end
