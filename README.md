# Asian Option Price Calculator

This project prices an Asian call option using:

- an Excel-based Monte Carlo simulation
- a VBA Monte Carlo engine
- a geometric Asian closed-form benchmark
- a convergence table for the number of simulations

## Files

- `workbook/options_project.xlsm`: Excel workbook with inputs, outputs and VBA macros
- `vba/`: exported VBA modules for code review
- `README.md`: project description

## Main inputs

- Initial stock price
- Strike
- Risk-free rate
- Dividend yield
- Volatility
- Maturity
- Number of time steps
- Number of simulations

## VBA macros

- `Run_Asian_Project`
- `Fill_Convergence_Table`
