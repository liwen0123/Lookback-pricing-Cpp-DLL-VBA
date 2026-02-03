# Lookback-pricing-Cpp-DLL-VBA
This project implements a numerical solver for pricing European-style Lookback Options using Monte-Carlo simulation within the Black-Scholes framework. It was developed as part of the M2 Quantitative Finance program.

The project consists of:
- A **C++ implementation** using Visual Studio.
- A **VBA-based user interface** integrated into Excel.
- A **64-bit DLL** compiled from C++, which is invoked directly by the VBA frontend.
# Project Overview
A Lookback option's payoff depends on the maximum or minimum price of the underlying asset over the period between emission and maturity.
- A **Lookback call**: Pays the difference between the underlying price $S(T)$ at maturity and its minimum price during the life of the option: $S(T) - \min_{T_0 \le t \le T} S(t)$.
- A**Lookback Put**: Pays the difference between the maximum price and the price at maturity: $\max_{T_0 \le t \le T} S(t) - S(T)$.

# Input Parameters
The following inputs are required for the simulation:
- Calculation Date: Default is the current date.
- Maturity: The expiration date of the option.
- Option Type: Selection between Call or Put.
- Spot Price ($S_0$): Current price of the underlying asset.
- Risk-Free Rate ($r$): Constant interest rate.
- Volatility ($\sigma$).
- Monte-Carlo Parameters: Number of simulations and time steps
# Features
The program calculates the following outputs:
- Theoretical Price ($P$).
- Greeks: $\Delta$ (Delta), $\Gamma$ (Gamma), $\Theta$ (Theta), $\rho$ (Rho), and $Vega$.
- Visualization:
   - Graph of Option Price $P(S, T_0)$ vs. Underlying Price $S_0$.
   - Graph of Delta $\Delta(S, T_0)$ vs. Underlying Price $S_0$.
