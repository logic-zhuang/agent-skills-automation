set more off
sysuse auto, clear
regress mpg weight foreign
predict r, residuals
scatter r weight, yline(0) title("Residuals vs weight")
