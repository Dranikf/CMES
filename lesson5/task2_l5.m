function b = task2_l5(a, s)
    b = [];
    for i = 1:3
        b = [b , normrnd(a, s)];
    end

end