## Convert MUMPS dates to Excel

![image info](./toolbar.png)

Some Epic™ systems report dates in [MUMPS format](https://en.wikipedia.org/wiki/MUMPS), which stores dates/times as either the number of seconds (or days) since 31 December 1840[^1]. These need to be converted to Excel format (fractional days since 01 January 1900) if we are to make sense of them.

Here's an example of MUMPS data in elapsed seconds:

![image info](./mumps_dates.png)

...and the same data in elapsed *days*:

![image info](./mumps_dates_in_days.png)

...and the converted values:

![image info](./converted.png)

The app figures out whether the MUMPS data are in seconds or days based on the size of the numbers.

[BACK](../../README.md)

[^1]: https://ken-blog.krugler.org/2011/07/15/interesting-dates-in-computer-programming/