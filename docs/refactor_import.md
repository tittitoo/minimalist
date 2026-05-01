# Refactor import

Some of the imports are not needed for every function calls such as
`reportlab` and `requests`. We should refactor these into conditional
import.

Let us make a plan on this.
