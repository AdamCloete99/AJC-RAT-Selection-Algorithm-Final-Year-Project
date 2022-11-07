%Aggregation

filename = 'supplierSelection.xlsx';
sheetread = 2;
nDMrange = 'c22:f25';
%nDM = xlsread(filename,sheetread, nDMrange);
nDM = readmatrix(filename,'Sheet',sheetread, 'Range', nDMrange);

nVoice_range = 'AA3:AD103';
nVideo_range = 'AE3:AH103';
nData_range = 'AI3:AL103';

%nVoice = xlsread(filename,sheetread, nVoice_range);
nVoice = readmatrix(filename,'Sheet',sheetread, 'Range', nVoice_range);

%nVideo = xlsread(filename,sheetread, nVideo_range);
nVideo = readmatrix(filename,'Sheet',sheetread, 'Range', nVideo_range);

%nData = xlsread(filename,sheetread, nData_range);
nData = readmatrix(filename,'Sheet',sheetread, 'Range', nData_range);

user = 1:100;
weight = 1:4;
for i = 1:length(user)
    for j = 1:length(weight)
        for k = 1:length(weight)
            aDM(j,k) = nDM(j,k) * nVoice(i,k); %only voice
            
        end
    end 
    MX = max(aDM);
    MN = min(aDM);
    FPIS = [MN(1), MX(2), MN(3), MX(4)];
    FNIS = [MX(1), MN(2), MX(3), MN(4)];

    for l = 1:length(weight)
        for m = 1:length(weight)
            distPsq(m) = (aDM(l,m) - FPIS(m))^2;
            distNsq(m) = (aDM(l,m) - FNIS(m))^2;
            
        end
        distP = sqrt(sum(distPsq));
        distN = sqrt(sum(distNsq));
        
        CC(l) = distN/(distN+distP);
    end
    [val, pos] = max(CC);
    selVoice(i) = pos;
    CC(pos) = 0;
    [val, post] = max(CC);
    selVoice_two(i) = post;
    CC(post) = 0;
    [val, posl] = max(CC);
    selVoice_three(i) = posl;
end
writematrix(selVoice, 'SupplierSelection.xlsx', 'Sheet', 3, 'Range', 'B2:CW2')
writematrix(selVoice_two, 'SupplierSelection.xlsx', 'Sheet', 3, 'Range', 'B6:CW6')
writematrix(selVoice_three, 'SupplierSelection.xlsx', 'Sheet', 3, 'Range', 'B10:CW10')

%video
for i = 1:length(user)
    for j = 1:length(weight)
        for k = 1:length(weight)
            aDM(j,k) = nDM(j,k) * nVideo(i,k); %only video
            
        end
    end 
    MX = max(aDM);
    MN = min(aDM);
    FPIS = [MN(1), MX(2), MN(3), MX(4)];
    FNIS = [MX(1), MN(2), MX(3), MN(4)];

    for l = 1:length(weight)
        for m = 1:length(weight)
            distPsq(m) = (aDM(l,m) - FPIS(m))^2;
            distNsq(m) = (aDM(l,m) - FNIS(m))^2;
            
        end
        distP = sqrt(sum(distPsq));
        distN = sqrt(sum(distNsq));
        
        CC(l) = distN/(distN+distP);
    end
    [val, pos] = max(CC);
    selVideo(i) = pos;
    CC(pos) = 0;
    [val, post] = max(CC);
    selVideo_two(i) = post;
    CC(post) = 0;
    [val, posl] = max(CC);
    selVideo_three(i) = posl;
end
writematrix(selVideo, 'SupplierSelection.xlsx', 'Sheet', 3, 'Range', 'B3:CW3')
writematrix(selVideo_two, 'SupplierSelection.xlsx', 'Sheet', 3, 'Range', 'B7:CW7')
writematrix(selVideo_three, 'SupplierSelection.xlsx', 'Sheet', 3, 'Range', 'B11:CW11')

%data
for i = 1:length(user)
    for j = 1:length(weight)
        for k = 1:length(weight)
            aDM(j,k) = nDM(j,k) * nData(i,k); %only data
            
        end
    end 
    MX = max(aDM);
    MN = min(aDM);
    FPIS = [MN(1), MX(2), MN(3), MX(4)];
    FNIS = [MX(1), MN(2), MX(3), MN(4)];

    for l = 1:length(weight)
        for m = 1:length(weight)
            distPsq(m) = (aDM(l,m) - FPIS(m))^2;
            distNsq(m) = (aDM(l,m) - FNIS(m))^2;
            
        end
        distP = sqrt(sum(distPsq));
        distN = sqrt(sum(distNsq));
        
        CC(l) = distN/(distN+distP);
    end
    [val, pos] = max(CC);
    selData(i) = pos;
    CC(pos) = 0;
    [val, post] = max(CC);
    selData_two(i) = post;
    CC(post) = 0;
    [val, posl] = max(CC);
    selData_three(i) = posl;
end
writematrix(selData, 'SupplierSelection.xlsx', 'Sheet', 3, 'Range', 'B4:CW4')
writematrix(selData_two, 'SupplierSelection.xlsx', 'Sheet', 3, 'Range', 'B8:CW8')
writematrix(selData_three, 'SupplierSelection.xlsx', 'Sheet', 3, 'Range', 'B12:CW12')