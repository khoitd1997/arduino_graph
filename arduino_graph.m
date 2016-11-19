se=instrhwinfo('serial');%find the available port on tbe computer, run again if the usb is disconnected
d=[];
if isempty(se.AvailableSerialPorts())
    return
else
    hold on
    a=se.AvailableSerialPorts();
    port=serial(a{1});
    fopen(port);
    readasync(port);
    see=instrhwinfo('serial');
    while isempty(see.AvailableSerialPorts())~=0
    a=fscanf(port);
    b=textscan(a,'%d');
    d=[d;b{1}(1) b{1}(2)];%build a matrix of data
    app.NumericEditField2.Value=b{1}(2);
    plot(b{1}(1),b{1}(2),'.k');
    pause(55/1000)
    see=instrhwinfo('serial');
    end
    stopasync(port);
    fclose(port);
    clc
    clear a b se
    clear instrfind
    disp(d)
    hold off
    %fclose(instrfind) will mostly fix any problems
    %https://www.mathworks.com/matlabcentral/answers/252467-xlswrite-function-is-continuosly-providing-error-error-using-xlswrite-line-219-invoke-error
end
