# Prevent Excel from Stealing Focus

When running functions like `conditional_format_wb` the sheets are
activated one by one for the corresponding micros to be run. The
conditional formatting is handled by excel VBA.

The undesirable affect of this is that when a sheet is activated, focus is
shifted to the excel app, which prevents the user from being able to do
other things, like email.

Can we plan to have the excel run silently without stealing the focus? You
may ask me for clarification.
