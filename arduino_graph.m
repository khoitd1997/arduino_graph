%% Initial setup
se=instrhwinfo('serial');%find the available port on tbe computer, run again if the usb is disconnected
%d=[]; %empty array slows down data collection 
d=ones(10000,2); %capable of storing 10000 run change as you need, limited only by ram of the computer
count=1;
if isempty(se.AvailableSerialPorts())
    return
else
    hold on
    a=se.AvailableSerialPorts();
    port=serial(a{1});
    fopen(port);
    readasync(port);
    see=instrhwinfo('serial');
%% loop start here, loop until the arduino is unplugged(try bluetooth see if it works)
    while isempty(see.SerialPorts)~=1%check if there is any arduino plugged in
    a=fscanf(port);
    b=textscan(a,'%d');% %d is working fine but not for integer, try %f
    
    %build a matrix of data
    d(count,1)=b{1}(1);
    d(count,2)=b{1}(2);
    
    plot(b{1}(1),b{1}(2),'-*k');
    pause(20/1000)
    see=instrhwinfo('serial');
    count=count+1;
    end
%% cleaning up
    stopasync(port);
    fclose(port);
    file_name=[strrep(datestr(now),':','`'),'.xlsm']; 
    d(count:end,:)=[];%truncating the matrix from row number count to end
    xlswrite(file_name,d);%work but a little slow, tested with a small maxtrix, about the same time as writetable but no header
    clc
    clear a b se count d 
    clear instrfind
    hold off
    %fclose(instrfind) will mostly fix any problems, use when u r doomed
    %https://www.mathworks.com/matlabcentral/answers/252467-xlswrite-function-is-continuosly-providing-error-error-using-xlswrite-line-219-invoke-error
end
