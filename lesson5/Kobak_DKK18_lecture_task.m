val = normrnd(0,1, 20, 1);

%построение первого графического окна
subplot(1, 2, 1);
hist(val);
x = -3:0.1:3;
hold on;
plot(x, normpdf(x,0,1), 'LineWidth', 5);
subplot(1 ,2 ,1);

%построение второго графического окна
subplot(1, 2, 2);
plot(x, normcdf(x , 0, 1));
hold on
val = sort(val);
stairs(val, normcdf(val, 0, 1));
legend('teory function', 'empirical fuction')