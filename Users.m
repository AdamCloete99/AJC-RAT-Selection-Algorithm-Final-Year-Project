%Generating 100 users with individual user preference weights 

%class = 1:3;
user = 1:100;
weight = 1:4;
for i = 1:length(user)
    for j = 1:length(weight)
%         %General:
%         Voice(i,j) = randi(9);
%         Video(i,j) = randi(9);
%         Data(i,j) = randi(9);

        %cost==1, data rate==2, delay==3, security==4
        if j==2
            Voice(i,j) = 5+randi(4);
            Video(i,j) = 5+randi(4);
            Data(i,j) = 5+randi(4);
        elseif j==3
            Voice(i,j) = 5+randi(4);
            Video(i,j) = 5+randi(4);
            Data(i,j) = 5+randi(4);

        else
            Voice(i,j) = randi(5);
            Video(i,j) = randi(5);
            Data(i,j) = randi(5);
        end
    end
end 
%xlswrite( 'SupplierSelection.xlsx',Voice,'InputData' ,'L3:N13') %L3:N13
writematrix(Voice, 'SupplierSelection.xlsx', 'Sheet', 2, 'Range', 'L3:o103')
writematrix(Video, 'SupplierSelection.xlsx', 'Sheet', 2, 'Range', 'P3:S103')
writematrix(Data, 'SupplierSelection.xlsx', 'Sheet', 2, 'Range', 'T3:W103')