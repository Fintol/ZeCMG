
function [ r ] = ZOS_API_matlab_AR_diffractive_waveguide_source( args )

if ~exist('args', 'var')
    args = [];
end

% Initialize the OpticStudio connection 初始化建立ZOS链接
TheApplication = InitConnection();
if isempty(TheApplication)
    % failed to initialize a connection
    r = [];
else
    try
        %运行主程序：新建或打开文件、建模、修改参数、光线追迹、数据分析、保存文件等
        r = BeginApplication(TheApplication, args);
        CleanupConnection(TheApplication);
    catch err
        CleanupConnection(TheApplication);
        rethrow(err);
    end
end
end

function [r] = BeginApplication(TheApplication, args)

import ZOSAPI.*;

    % creates a new API directory  
    % 新建一个API文件路径
    apiPath = System.String.Concat(TheApplication.SamplesDir, '\API\Matlab');
    if (exist(char(apiPath)) == 0) mkdir(char(apiPath)); end
    
    % Set up primary optical system
    % 将zemax光学系统命名为简单matlab变量：TheSystem，
    % 后续使用TheSystem.xxx 命令可直接读取或修改系统相应参数
    % xxxDir 提取部分系统默认路径供后续使用
    TheSystem = TheApplication.PrimarySystem;
    sampleDir = TheApplication.SamplesDir;
    ZemaxDir = TheApplication.ZemaxDataDir;
    
    % (可选) 自定义入射光线位置、方向角余弦、能量，保存为zemax source file文件，作为后续自定义光源使用    
    ray_xyz_lmn_I_wav(1,:)= [2,2,3,0.707,0,-0.707,1, 0.611 ]; % 1~7列数据依次为x,y,z,L,M,N,Intensity
    ray_xyz_lmn_I_wav(2,:)= [0,0,3,0.707,0,-0.707,1, 0.532 ]; % 第8列数据为波长，单位um，第8列可省略
    ray_count=size(ray_xyz_lmn_I_wav,1); %自定义光线总数目 
    zemax_source_file_dimension_flag=4; % zemax长度单位，4对应mm
    %打开或新建source文件，覆盖已有内容
    source_file_dat=fopen([char(ZemaxDir),'\Objects\Sources\Source Files\customized_source_file.dat'],'w');
    %写入source文件第一行
    fprintf(source_file_dat, '%u\t%u\n', ray_count, zemax_source_file_dimension_flag); 
    %写入所有光线信息至dat文件
    fprintf(source_file_dat, '%10.6f\t%10.6f\t%10.6f\t%10.6f\t%10.6f\t%10.6f\t%10.6f\t%10.6f\n', ray_xyz_lmn_I_wav'); 
    fclose(source_file_dat);

    %! [e02s01_m][e01s01_m] 代码来自Zemax官方案例e02s01 e01s01   
    % Make new file 新建zemax空文件
    newFile = System.String.Concat(sampleDir, '\API\Matlab\AR_waveguide_new.zos');
    TheSystem.New(false);
    TheSystem.MakeNonSequential(); %将zemax文件切换为非序列模式
    TheSystem.SystemData.MaterialCatalogs.AddCatalog('SCHOTT');
    TheSystem.SaveAs(newFile); %以对应路径和文件名保存zemax文件

    %在system explorer中添加波长
    wavelength_to_add= 0.456; wavelength_weight=1;
    TheSystem.SystemData.Wavelengths.AddWavelength(wavelength_to_add,wavelength_weight)

    %在system explorer中修改非序列模式参数，根据需求修改光线segment数、阈值、系统长度单位、光源单位等
    min_rel_ray_intensity= 1e-6; %追迹截止相对能量阈值
    TheSystem.SystemData.NonsequentialData.MinimumRelativeRayIntensity=min_rel_ray_intensity;

    
    %非序列模式editor，后续使用TheNCE直接修改editor内参数
    TheNCE = TheSystem.NCE; 
   
    % 在指定编号位置插入新object，使用eval()方便后续批量添加/修改object
    NCE_object_index=1; %非序列模式element 序号
    % obj_1=TheNCE.InsertNewObjectAt(NCE_object_index); %与下行执行动作相同
    eval( ['obj_', num2str(NCE_object_index), '=TheNCE.InsertNewObjectAt(NCE_object_index);' ] );

    % 获取object类型SourceFile
    object_type=eval( ['obj_', num2str(NCE_object_index), '.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.SourceFile);' ] );
    % 指定SourceFile使用光线文件名
    object_type.FileName1='customized_source_file.dat';
    % 将object 1 的类型修改为SourceFile
    eval( ['obj_', num2str(NCE_object_index), '.ChangeType(object_type);' ] );

    % 修改光线的偏振态
    zemax_jx=1; zemax_jy=1; zemax_xphase=90;
    eval( ['obj_', num2str(NCE_object_index), '.SourcesData.Jx=zemax_jx;' ] );
    eval( ['obj_', num2str(NCE_object_index), '.SourcesData.Jy=zemax_jy;' ] );
    eval( ['obj_', num2str(NCE_object_index), '.SourcesData.XPhase=zemax_xphase;' ] );

    % 修改object位置
    xyz=[0,0,0];
    eval( ['obj_', num2str(NCE_object_index), '.XPosition=xyz(1);' ] );
    eval( ['obj_', num2str(NCE_object_index), '.YPosition=xyz(2);' ] );
    eval( ['obj_', num2str(NCE_object_index), '.ZPosition=xyz(3);' ] );

    % 修改layout光线数和analysis光线数
    layout_ray=1; analysis_ray=1e6;
    eval( ['obj_', num2str(NCE_object_index), '.ObjectData.NumberOfAnalysisRays=analysis_ray;' ] );
    eval( ['obj_', num2str(NCE_object_index), '.ObjectData.NumberOfLayoutRays=layout_ray;' ] );

    % 在layout中不绘制光源
    eval( ['obj_', num2str(NCE_object_index), '.DrawData.DoNotDrawObject=1;' ] );
    
    %保存文件
    TheSystem.Save();

    r = [];
end

function app = InitConnection()

import System.Reflection.*;

% Find the installed version of OpticStudio.
zemaxData = winqueryreg('HKEY_CURRENT_USER', 'Software\Zemax', 'ZemaxRoot');
NetHelper = strcat(zemaxData, '\ZOS-API\Libraries\ZOSAPI_NetHelper.dll');
% Note -- uncomment the following line to use a custom NetHelper path
% NetHelper = 'C:\Users\Documents\Zemax\ZOS-API\Libraries\ZOSAPI_NetHelper.dll';
NET.addAssembly(NetHelper);

success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
% Note -- uncomment the following line to use a custom initialization path
% success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize('C:\Program Files\OpticStudio\');
if success == 1
    LogMessage(strcat('Found OpticStudio at: ', char(ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory())));
else
    app = [];
    return;
end

% Now load the ZOS-API assemblies
NET.addAssembly(AssemblyName('ZOSAPI_Interfaces'));
NET.addAssembly(AssemblyName('ZOSAPI'));

% Create the initial connection class
TheConnection = ZOSAPI.ZOSAPI_Connection();

% Attempt to create a Standalone connection

% NOTE - if this fails with a message like 'Unable to load one or more of
% the requested types', it is usually caused by try to connect to a 32-bit
% version of OpticStudio from a 64-bit version of MATLAB (or vice-versa).
% This is an issue with how MATLAB interfaces with .NET, and the only
% current workaround is to use 32- or 64-bit versions of both applications.
app = TheConnection.CreateNewApplication();
if isempty(app)
   HandleError('An unknown connection error occurred!');
end
if ~app.IsValidLicenseForAPI
    HandleError('License check failed!');
    app = [];
end

end

function LogMessage(msg)
disp(msg);
end

function HandleError(error)
ME = MXException(error);
throw(ME);
end

function  CleanupConnection(TheApplication)
% Note - this will close down the connection.

% If you want to keep the application open, you should skip this step
% and store the instance somewhere instead.
TheApplication.CloseApplication();
end




