## Hide Dr. Names

![image info](./toolbar.png)

In some cases, we want to study patient or physician notes but not reveal the physicians' names. Using this tool, we can search for physician names & replace with a scrambled identifier. If the notes look like this:

![image info](./raw_data.png)

...clicking the `Hide Dr. Names` button starts the search. If a string matches one of the app's regular expressions, a GUI asks the user for confirmation:

![image info](./gui.png)

If the detected name seems like one previously identified, the GUI allows the user to __link__ the two versions:

![image info](./link.png)

The processed notes look like this. Note that linking `Provider, Able A., MD` with `Dr. Able Provider` means the same scrambled identifier is used in the two notes. This enables notes mentioning the same physician can be studied as a group.

![image info](./converted.png)

[BACK](../../README.md)