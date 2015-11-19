# LFEG - Lighitng Fast Excel Generator

The library was built to generate big excel reports which don't need a lot of fancy formatting or any formulas. It's similar to CSV files but Excel sometimes doesn't correctly determines data types for columns and because of that sorting, aggregate and other functions don't work out of the box. 

There are .net libraries which allow to create Excel files but the ones I saw were very-very slow when you have a lot of data. But this library allows to generate files with 5000 rows in 200ms or so depending on hardware. And even 1000000 rows reports are not a problem, it will take less than half of a minute to generate.
