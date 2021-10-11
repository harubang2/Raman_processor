function varargout = Raman_pre_processing(varargin)
% RAMAN_PRE_PROCESSING MATLAB code for Raman_pre_processing.fig
%      RAMAN_PRE_PROCESSING, by itself, creates a new RAMAN_PRE_PROCESSING or raises the existing
%      singleton*.
%
%      H = RAMAN_PRE_PROCESSING returns the handle to a new RAMAN_PRE_PROCESSING or the handle to
%      the existing singleton*.
%
%      RAMAN_PRE_PROCESSING('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in RAMAN_PRE_PROCESSING.M with the given input arguments.
%
%      RAMAN_PRE_PROCESSING('Property','Value',...) creates a new RAMAN_PRE_PROCESSING or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before Raman_pre_processing_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to Raman_pre_processing_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help Raman_pre_processing

% Last Modified by GUIDE v2.5 18-Mar-2021 11:45:55

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Raman_pre_processing_OpeningFcn, ...
                   'gui_OutputFcn',  @Raman_pre_processing_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before Raman_pre_processing is made visible.
function Raman_pre_processing_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to Raman_pre_processing (see VARARGIN)

% Choose default command line output for Raman_pre_processing
handles.output = hObject;

global upper lower Poly_order_SG Frame_window ORDER THRESHOLD FCT

axes(handles.axes_1);
imshow('Logo.jpg');

axes(handles.axes_2);
imshow('Logo.jpg');

axes(handles.axes_3);
imshow('Logo.jpg');

lower = 400;            % lower limit of spectral region of interest
upper = 3300;           % upper limit of spectral region of interest

Poly_order_SG = 1;      % polynomial order for Savitzky-Golay filter
Frame_window = 15;      % Frame window for Savitzky-Golat filter

ORDER = 2;              % polynomial order for baseline subtraction
THRESHOLD = 0.01;
FCT = 'atq';

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Raman_pre_processing wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Raman_pre_processing_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in SG_filter.
function SG_filter_Callback(hObject, eventdata, handles)
global upper lower x y_raw y_mid Poly_order_SG Frame_window cell_interest ret y

ret = [];
y = [];
y_mid = [];

y_mid(:,cell_interest) = sgolayfilt(y_raw(:,cell_interest), Poly_order_SG, Frame_window);

axes(handles.axes_2);
plot(x, y_mid(:,cell_interest), 'b-', 'LineWidth', 1)
% title('Measured spectra')
xlabel('Raman shift (cm^-^1)')
ylabel('Intensity (AU)')
axis([lower upper min(y_raw(:,cell_interest)) max(y_raw(:,cell_interest))])
set(gca,'fontsize',8)
grid on

% hObject    handle to SG_filter (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% --- Executes on button press in Base_subtraction.
function Base_subtraction_Callback(hObject, eventdata, handles)
global upper lower x y y_raw y_mid cell_interest ORDER THRESHOLD FCT

[EST,COEFS,IT,ORDER,THRESHOLD,FCT] = backcor(x,y_mid(:,cell_interest));

y(:,cell_interest) = y_mid(:,cell_interest)-EST;
iteration = IT;

axes(handles.axes_2);
plot(x, y_mid(:,cell_interest), 'b-', 'LineWidth', 1)
% title('Measured spectra')
xlabel('Raman shift (cm^-^1)')
ylabel('Intensity (AU)')
axis([lower upper min(y_raw(:,cell_interest)) max(y_raw(:,cell_interest))])
set(gca,'fontsize',8)
grid on
hold on
plot(x, EST, 'm--', 'LineWidth', 0.5)
set(gca,'fontsize',8)
hold off

axes(handles.axes_3);
plot(x, y(:,cell_interest), 'r-', 'LineWidth', 1)
legend(sprintf('Number of iteration = %d',iteration))
% title('Measured spectra')
xlabel('Raman shift (cm^-^1)')
ylabel('Intensity (AU)')
axis([lower upper -inf inf])
set(gca,'fontsize',8)
grid on

% hObject    handle to Base_subtraction (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% --- Executes on button press in Import.
function Import_Callback(hObject, eventdata, handles)
% hObject    handle to Import (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global upper lower x y_raw M

set(handles.List_1,'String',[]);

filename = 'input.xlsx';
M_raw = xlsread(filename);
M = M_raw.';
y_raw = [];
x = [];
y_ini = [];

bar_2 = waitbar(0,'Importing...');

for K = 1 : size(M,2)-1
    waitbar(K/(size(M,2)-1),bar_2)
    
    if K == 1
        x = M(:,1);
    end
    y_ini = M(:,K+1);
        
    % select the spectral region of interest for processing
    y_ini(x > upper | x < lower) = [];
       
    for J = 1 : size(y_ini,1)
        y_raw(J,K) = y_ini(J);
    end
    
    old_str = get(handles.List_1,'String');
    new_str = strvcat(old_str, num2str(K));
    set(handles.List_1,'String',new_str);
end
close(bar_2)
x(x > upper | x < lower) = [];

function Poly_order_SG_Callback(hObject, eventdata, handles)
% hObject    handle to Poly_order_SG (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global Poly_order_SG

Poly_order_SG = str2double(get(handles.Poly_order_SG,'String'));

% Hints: get(hObject,'String') returns contents of Poly_order_SG as text
%        str2double(get(hObject,'String')) returns contents of Poly_order_SG as a double

% --- Executes during object creation, after setting all properties.
function Poly_order_SG_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Poly_order_SG (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


function Frame_window_Callback(hObject, eventdata, handles)
% hObject    handle to Frame_window (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global Frame_window

Frame_window = str2double(get(handles.Frame_window,'String'));

% Hints: get(hObject,'String') returns contents of Frame_window as text
%        str2double(get(hObject,'String')) returns contents of Frame_window as a double


% --- Executes during object creation, after setting all properties.
function Frame_window_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Frame_window (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% function Poly_order_fit_Callback(hObject, eventdata, handles)
% % hObject    handle to Poly_order_fit (see GCBO)
% % eventdata  reserved - to be defined in a future version of MATLAB
% % handles    structure with handles and user data (see GUIDATA)
% global Poly_order_fit
% 
% Poly_order_fit = str2double(get(handles.Poly_order_fit,'String'));

% Hints: get(hObject,'String') returns contents of Poly_order_fit as text
%        str2double(get(hObject,'String')) returns contents of Poly_order_fit as a double


% --- Executes during object creation, after setting all properties.
% function Poly_order_fit_CreateFcn(hObject, eventdata, handles)
% % hObject    handle to Poly_order_fit (see GCBO)
% % eventdata  reserved - to be defined in a future version of MATLAB
% % handles    empty - handles not created until after all CreateFcns called
% 
% % Hint: edit controls usually have a white background on Windows.
% %       See ISPC and COMPUTER.
% if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
%     set(hObject,'BackgroundColor','white');
% end


function Lower_Callback(hObject, eventdata, handles)
% hObject    handle to Lower (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global lower

lower = str2double(get(handles.Lower,'String'));

% axes(handles.axes_1);
% axis([lower upper -inf inf])

% Hints: get(hObject,'String') returns contents of Lower as text
%        str2double(get(hObject,'String')) returns contents of Lower as a double


% --- Executes during object creation, after setting all properties.
function Lower_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Lower (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


function Upper_Callback(hObject, eventdata, handles)
% hObject    handle to Upper (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global upper

upper = str2double(get(handles.Upper,'String'));

% axes(handles.axes_1);
% axis([lower upper -inf inf])

% Hints: get(hObject,'String') returns contents of Upper as text
%        str2double(get(hObject,'String')) returns contents of Upper as a double


% --- Executes during object creation, after setting all properties.
function Upper_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Upper (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in List_1.
function List_1_Callback(hObject, eventdata, handles)
% hObject    handle to List_1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global lower upper x y_raw cell_interest

cell_interest = get(handles.List_1,'Value');
axes(handles.axes_1);
plot(x, y_raw(:,cell_interest), 'k-', 'LineWidth', 1)
legend(sprintf('Cell# %d',cell_interest))
% title('Measured spectra')
xlabel('Raman shift (cm^-^1)')
ylabel('Intensity (AU)')
axis([lower upper min(y_raw(:,cell_interest)) max(y_raw(:,cell_interest))])
set(gca,'fontsize',8)
grid on

% Hints: contents = cellstr(get(hObject,'String')) returns List_1 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from List_1


% --- Executes during object creation, after setting all properties.
function List_1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to List_1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in Export.
function Export_Callback(hObject, eventdata, handles)
global x y cell_interest Poly_order_SG Frame_window ORDER THRESHOLD FCT

cell_number = ['cell# ' num2str(cell_interest)];
output2 = ['Summary_cell_' num2str(cell_interest) '.xlsx'];
A1 = {'S-G_polynomial order'; 'S-G_Framelength'; 'Baseline_polynomial order'; 'Baseline_threshold'; 'Baseline_function'};
T1 = [Poly_order_SG, Frame_window, ORDER, THRESHOLD, convertCharsToStrings(FCT)]';
A2 = {'Raman shift (cm-1)'; cell_number};
T2 = [x, y(:,cell_interest)]';
sheet = 1;
xlRange = 'B1';
xlswrite(output2, A1, sheet)
xlswrite(output2, T1, sheet, xlRange)
xlRange_2 = 'A6';
xlswrite(output2, A2, sheet, xlRange_2)
xlRange_3 = 'B6';
xlswrite(output2, T2, sheet, xlRange_3)


% --- Executes on button press in Batch_proc.
function Batch_proc_Callback(hObject, eventdata, handles)
global M upper lower loopFlag ORDER THRESHOLD FCT

loopFlag = 0;

set(handles.List_2,'String', []);

for K = 1 : size(M,2)-1
    x = M(:,1);
    y_raw = M(:,K+1);
    
    % choose the spectral region for the preprocessing
    y_raw(x > upper) = [];           % upper bound
    y_raw(x < lower) = [];           % lower bound
    x(x > upper) = [];
    x(x < lower) = [];
        
    y_mid = SG(y_raw);                                % Savitzky-Golay filter
%     [y, ret, new_str] = MP(x, y_mid, K);              % I-Mod-Poly fitting
    [EST,COEFS,IT] = backcor(x,y_mid,ORDER,THRESHOLD,FCT);  % Baseline subtraction
    
    y = y_mid-EST;
    old_str = get(handles.List_2,'String');
    result = ['Cell# ' num2str(K)];
    result_2 = ['Number of iteration = ' num2str(IT)];
    new_str = strvcat(result, result_2);
    final_str = strvcat(old_str, new_str);
    if K == size(M,2)-1
        goodbye = ['Batch processing ends'];
        final_str = strvcat(old_str, new_str, goodbye);
    end
    set(handles.List_2,'String', final_str);
    index = size(get(handles.List_2,'String'), 1);
    set(handles.List_2,'Value', index);
    pause(0.01)
    Output(K, x, y_raw, y_mid, y);               % Write a modified data
    if loopFlag == 1;
        break
    end
end

% hObject    handle to Batch_proc (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

function [y_mid] = SG(y_raw)
global Poly_order_SG Frame_window

y_mid = sgolayfilt(y_raw, Poly_order_SG, Frame_window);


function [] = Output(K, x, y_raw, y_mid, y)
global Poly_order_SG Frame_window ORDER THRESHOLD FCT

name = ['cell# ' num2str(K)];

% Output: write a summary with modified spectra only
if K == 1
    output2 = ['Summary.xlsx'];
    A1 = {'S-G_polynomial order'; 'S-G_Framelength'; 'Baseline_polynomial order'; 'Baseline_threshold'; 'Baseline_function'};
    T1 = [Poly_order_SG, Frame_window, ORDER, THRESHOLD, convertCharsToStrings(FCT)]';
    A2 = {'Raman shift (cm-1)'; name};
    T2 = [x, y]';
    sheet = 1;
    xlRange = 'B1';
    xlswrite(output2, A1, sheet)
    xlswrite(output2, T1, sheet, xlRange)
    xlRange_2 = 'A6';
    xlswrite(output2, A2, sheet, xlRange_2)
    xlRange_3 = 'B6';
    xlswrite(output2, T2, sheet, xlRange_3)
else
    output2 = ['Summary.xlsx'];
    A2 = {name};
    T2 = [y];
    sheet = 1;
    xlRange1 = sprintf('A%d',K+6);
    xlRange2 = sprintf('B%d',K+6);
    xlswrite(output2, A2, sheet, xlRange1)
    xlswrite(output2, T2', sheet, xlRange2)
end


% --- Executes on selection change in List_2.
function List_2_Callback(hObject, eventdata, handles)
% hObject    handle to List_2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns List_2 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from List_2


% --- Executes during object creation, after setting all properties.
function List_2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to List_2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton6.
function pushbutton6_Callback(hObject, eventdata, handles)
global loopFlag

loopFlag = 1;
% hObject    handle to pushbutton6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in Export_mid.
function Export_mid_Callback(hObject, eventdata, handles)
global x y_mid cell_interest Poly_order_SG Frame_window

cell_number = ['cell# ' num2str(cell_interest)];
output3 = ['S-G_only_cell_' num2str(cell_interest) '.xlsx'];
A1 = {'S-G_polynomial order'; 'S-G_Framelength'};
T1 = [Poly_order_SG, Frame_window]';
A2 = {'Raman shift (cm-1)'; cell_number};
T2 = [x, y_mid(:,cell_interest)]';
sheet = 1;
xlRange = 'B1';
xlswrite(output3, A1, sheet)
xlswrite(output3, T1, sheet, xlRange)
xlRange_2 = 'A3';
xlswrite(output3, A2, sheet, xlRange_2)
xlRange_3 = 'B3';
xlswrite(output3, T2, sheet, xlRange_3)

% hObject    handle to Export_mid (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% --- Executes on button press in pushbutton8.
function pushbutton8_Callback(hObject, eventdata, handles)
global M upper lower loopFlag

loopFlag = 0;

set(handles.List_2,'String', []);

for K = 1 : size(M,2)-1
    x = M(:,1);
    y_raw = M(:,K+1);
    
    % choose the spectral region for the preprocessing
    y_raw(x > upper) = [];           % upper bound
    y_raw(x < lower) = [];           % lower bound
    x(x > upper) = [];
    x(x < lower) = [];
        
    y_mid = SG(y_raw);                                % Savitzky-Golay filter
    result = ['Cell# ' num2str(K) '  processing...'];
    new_str = strvcat(result);
    old_str = get(handles.List_2,'String');
    final_str = strvcat(old_str, new_str);
    if K == size(M,2)-1
        goodbye = ['Batch processing ends'];
        final_str = strvcat(old_str, new_str, goodbye);
    end
    set(handles.List_2,'String', final_str);
    index = size(get(handles.List_2,'String'), 1);
    set(handles.List_2,'Value', index);
    pause(0.01)
    Output_SG(K, x, y_raw, y_mid);                    % Write a modified data
    if loopFlag == 1;
        break
    end
end
% hObject    handle to pushbutton8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

function [] = Output_SG(K, x, y_raw, y_mid)
global Poly_order_SG Frame_window
% Output: write a modified spectrum (individual file)
% [pathstr, name, extension] = fileparts(thisfilename);
name = ['cell# ' num2str(K)];
% output = [name,'_background_substracted.xlsx'];
% A = {'Raman shift (cm-1)', 'Intensity_raw', 'Intensity_SG_filtered', 'Calculated_baseline', 'Intensity_baseline_corrected'};
% T = [x, y_raw, y_mid, ret, y];
% sheet = 1;
% xlRange = 'A2';
% xlswrite(output, A, sheet)
% xlswrite(output, T, sheet, xlRange)

% Output: write a summary with modified spectra only
if K == 1
    output2 = ['Summary_S-G_only.xlsx'];
    A1 = {'S-G_polynomial order'; 'S-G_Framelength'};
    T1 = [Poly_order_SG, Frame_window]';
    A2 = {'Raman shift (cm-1)'; name};
    T2 = [x, y_mid]';
    sheet = 1;
    xlRange = 'B1';
    xlswrite(output2, A1, sheet)
    xlswrite(output2, T1, sheet, xlRange)
    xlRange_2 = 'A3';
    xlswrite(output2, A2, sheet, xlRange_2)
    xlRange_3 = 'B3';
    xlswrite(output2, T2, sheet, xlRange_3)
else
    output2 = ['Summary_S-G_only.xlsx'];
    A2 = {name};
    T2 = [y_mid];
    sheet = 1;
    xlRange1 = sprintf('A%d',K+3);
    xlRange2 = sprintf('B%d',K+3);
    xlswrite(output2, A2, sheet, xlRange1)
    xlswrite(output2, T2', sheet, xlRange2)
end
