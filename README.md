# Greeks Calculator in VBA

This project provides a set of VBA functions to calculate the Greeks for European options (Call and Put) under the Black-Scholes model.

## Features

- Delta, Gamma, Vega, Theta, Rho
- Works directly in Excel formulas
- Easy to import via `.bas` module

## How to Use

1. Open the VBA editor (`Alt + F11`)
2. Import the `GreeksCalculator.bas` module
3. Use functions in Excel:

```excel
=DeltaCalc(100, 100, 1, 0.05, 0.2, "call")
