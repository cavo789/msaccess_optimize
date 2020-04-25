# Modules

## Save modules in a compiled state

If the database contains modules, be sure in the production version, to compile the source code. You just need to open one module (no importance) and to click on the `Debug` menu then `Compile`.

![Compiled state](./images/compiled_state.png)

## Use Option Explicit statement

If you've modules or form's event coded in VBA, open every module / forms (press then `ALT-F11` to open the editor) and in the very first line of your code, type `Option Explicit` like f.i.:

```vbnet
Option Compare Database
Option Explicit

Sub DoSomething()
    ' your own code
End Sub
```

Option Explicit force the VB Engine to check that variables exists before starting to run the code and not every time a variable is accessed. This give a (small) little speed improvement.

*Option Explicit is more an excellent way of coding: variables should be declared before using them*

## Unload references

Take a look, in your code, if you're referencing too much external dependencies.

From within the VB Editor, click on the `Tools` menu then select `References` and pay attention to the first items in the list, the checked one. Do you really need them?

An easy way to answer to this question is: uncheck them and click on the `OK` button.

Click on the `Debug` menu and choose `Compile`. If you get compilation errors (and no errors before), go back in the `References` window and check the library back.

![Unload unused references](./images/unload_references.png)

Most of time, only two references are needed:

1. `Visual Basic For Applications` (always at the top)
2. `Microsoft Access 16.0 Object Library` (the second one) (Note: "16.0" is variable and depends on your MS Office installed version)
