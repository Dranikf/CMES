a = [];

for i = 1:3
    for j = 1:3
        a(i , j) = normrnd(2, 3);

    end
end

a
z = a(1, 1)