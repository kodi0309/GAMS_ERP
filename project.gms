Sets
    i 'Products' /1*8/
    l 'Production lines' /1*3/
    o 'Orders' /1*5/
    m 'Materials' /1*10/;

$ontext
OrderQuantity(o,i) 'Order quantities in kg'
        1       2       3       4       5       6       7       8
    1   10000   0       800     0       0       0       0       0
    2   0       12500   0       0       300     900     0       0
    3   500     0       0       8000    0       0       600     0
    4   0       400     15000   0       0       0       0       2500
    5   0       0       0       100     0       9000    1500    0;

ProductComposition(i,m) 'Material composition of each product'
    /1.1 0.4, 1.2 0.3, 1.3 0.3,
     2.1 0.5, 2.4 0.5,
     3.6 0.3, 3.7 0.4, 3.2 0.3,
     4.3 0.1, 4.8 0.4, 4.9 0.3, 4.10 0.2,
     5.4 0.4, 5.6 0.4, 5.10 0.2,
     6.1 0.5, 6.7 0.5,
     7.9 0.2, 7.2 0.4, 7.5 0.25, 7.3 0.15,
     8.10 0.2, 8.4 0.3, 8.6 0.5/
$offText

Parameters
OrderQuantity(o, i) 'Orders table'
$call gdxxrw input=Orders.xlsx output=orders.gdx par=orderQuantity rng=Orders!A1:I6 rdim=1 cdim=1
$gdxin orders.gdx
$load orderQuantity

ProductComposition(i,m) 'Material composition of each product';
$call gdxxrw input=Ingredients.xlsx output=ingredients.gdx par=productComposition rng=Composition!A1:I11 rdim=1 cdim=1
$gdxin ingredients.gdx
$load productComposition

Parameters
SetupTime(i) 'Changeover time for products in hours'
        /1 4, 2 3, 3 5, 4 2, 5 6, 6 3, 7 4, 8 5/
ProductionRate(i) 'Production rate in kg/hour'
        /1 280, 2 250, 3 200, 4 240, 5 160, 6 205, 7 215, 8 135/
EmployeeCostPerHour
        /50/
FixedCostPerHour
        /350/
MaterialCost(m) 'Material cost per kg'
        /1 4.15, 2 3.25, 3 7.20, 4 5.50, 5 4.30, 6 3.20, 7 5.60, 8 4.00, 9 6.20, 10 5.75/
        
SellingPrice(i) 'Selling price per kilogram of product i'
        /1 8.00, 2 1.50, 3 11.20, 4 11.30, 5 11.15, 6 9.00, 7 9.50, 8 12.00/;

Scalar
    TimeLimit /1000/;


Variables
    TotalProfit 'Total profit';

Binary Variables
    y(i,l) 'Is product produced on line'
    z(o) 'Order selected';

Positive Variables
    x(o,i,l) 'Quantity of product i for order o on line l';

Equations
    Objective 'Maximize total profit'
    TimeLimitPerLine(l) 'Ensure limit of production speed'
    OrderCompletion(o,i) 'Ensure order completion'
    LineUsage(o,i,l) 'Link production to line usage'
    Consistency(o) 'Consistency in order selection';

Objective ..
    TotalProfit =e=
        sum(o, z(o) * sum(i, OrderQuantity(o,i) * SellingPrice(i))) -
        sum((o,i,m,l), x(o,i,l) * ProductComposition(i,m) * MaterialCost(m)) -
        sum((i,l), y(i,l) * SetupTime(i) * EmployeeCostPerHour) -
        sum((o,i,l), (x(o,i,l) / ProductionRate(i)) * (FixedCostPerHour + EmployeeCostPerHour));

TimeLimitPerLine(l) ..
    sum((o,i), x(o,i,l) / ProductionRate(i)) + sum(i, y(i,l) * SetupTime(i)) =l= TimeLimit;

OrderCompletion(o,i) ..
    sum(l, x(o,i,l)) =e= OrderQuantity(o,i) * z(o);

LineUsage(o,i,l) ..
    x(o,i,l) =l= OrderQuantity(o,i) * y(i,l);

Consistency(o) ..
    z(o) =l= sum((i,l)$(OrderQuantity(o,i) > 0), y(i,l));

Model MaximizeProfit /all/;
Solve MaximizeProfit maximizing TotalProfit using MIP;

Display x.l, y.l, z.l, TotalProfit.l;
