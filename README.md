## What is this?
Just a simple excel function (an UDF, User Defined Function) which allows calculating the Boltzmann population of several energy levels (eg: conformers, spin states, etc).

Once installed, the function can be called just like any other excel function, by digiting *"=Function_Name"*; in this case, *"=Boltzmann_Pop("*

The code is mainly intended for use with computational chemistry packages, hence the **default energy units** are **Hartrees** (a.u.), but **kJ/mol**, **kcal/mol**, **eV** can be set easily.
The default temperature is **298.15** K, but other temperatures can be set easily.

## Installation
The function is contained in the Boltzmann_Pop.xlam file.

Exhaustive instructions on installing these files can be found online, eg:
https://www.automateexcel.com/vba/install-add-in


## Some examples
We calculated the energies of the three low-energy conformations of butane: two isoenergetic *gauche* and a *trans* conformer.
We want to know how populated those conformations are at room temperature.
The energies reported here are in Hartree, so the "energy" parameter can be omitted, as Hartrees are set as default energy units.
In order to calculate the population at 298.15 K, we start by typing the formula:
*=Boltzmann_Pop(*

![](image/Im2.png)

Note that, while typing, a **HELP function** is displayed. This can be executed if the command syntax is forgotten: the "Help" function displays the following informative window.

![](image/Im5.png)

By typing
*=Boltzmann_Pop(Range;;)*
which is equivalent to:
*=Boltzmann_Pop(Range;298.15;"Hartrees")*...

![](image/Im3.png)

...the populations of the energy levels in the input range is displayed.
The obtained values state that at room temperature (in the gas phase) 67% of butane molecules are in *Anti* conformation, and the remaining 33% is in two isoenergetic (ie symmetric, degenerate) *"Gauche"* conformations

![](image/Im4.png)


By typing in another cell the command
*=Boltzmann_Pop(Range;3000;"Hartrees")*
the same calculation is performed at a higher temperature (3000 K).
We see that at 3000 K, only 36% of butane molecules are in Anti conformation, and 64% of butane molecules are in Gauche conformations.

![](image/Im6.png)

![](image/Im7.png)


## Some notes
As you can easily guess from the source code, the main user-specified options are regarding the energy units.
A bit of typing freedom is tolerated:
- The keywords  "Hartree", "hartree", "Hartrees", "hartrees", "Eh", "eh", "au", "a.u." **set units to Hartrees**
- "kJ/mol", "kj/mol", "kj\mol", "kJ\mol", "kjmol", "KJ/mol" **set units to kJ/mol** 
- "kcal/mol", "Kcal/mol", "kcalmol", "kcal\mol", "Kcal\mol" **set units to kcal/mol**
- "ev", "eV", "EV"  **set units to eV**



*Written by Vincenzo Brancaccio, 2020.
 You are free to share, edit, copy, improve or break this code!*
