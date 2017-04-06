'Look for 1 non-zero date, unites with "-" and gives the first date that is equal to `D44`

=TEXT(
INDEX(MyDates;1;VERGLEICH(WAHR;INDEX(E44:BK44>0;);0));
"MM/JJ")
&" - "&
TEXT(
INDEX(MyDates;1;VERGLEICH(WAHR;INDEX(E44:BK44=D44;);0));
"MM/JJ")
