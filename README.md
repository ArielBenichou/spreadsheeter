# SpreadSheeter
This is a script to help my father be lazy in his job.

The program load a table with products and thier ids, each row is a client, each row is a different type of products.

The program loads a second execl table with ids and prices.

And then the program SPREADSHEET them - for each product the program multuply the quantity (in the first table) by the price (in the second table).

At the End the program add in the last row a SUM of how much money the client must pay.

## Useage
In the `dist` folder there is an `.exe` file use it if you want the program.

(If you want the code you know what to do.)

The program open up with a simple UI: 2 'Browse' button and one 'Finish' button.
What to do:
1. load the first table with the first button.
2. load the seconds table with the second button.
3. click finish!
4. in the same folder of the first table you have an "_out.xlsx" file.
5. DONE!

### The first table - Clients
The table MUST be structered like this:
| ID    | NAME        | PRICE |
| ----- | ----------- | ----- |
| 1111  | Cumcumbers  | 3.2   |
| 9589  | Helicopers  | 50000 |
| 3323  | Red Meat    | 32.1  |

AKA:  
1. FIRST ROW is the headers
2. FIRST COLUMN is the ids
3. THIRD COLUMN is the Prices

### The seconds table - Prices
The table MUST be structered like this:
| ID              | 1111        | 9589        | 3323        |
| -----           | ----------- | -----       | ---         |
| Product Name    | Cumcumbers  | Helicopers  | Red Meat    |
| Yossi           | 1.2         | 3           | 5.2         |
| Dr. Shipship    | 0.3         | 0           | 0           |
| Rebecca Greenred| 0           | 32          | 1.4         |

AKA:  
1. FIRST ROW is the ids - exclude the first column
2. SECOND ROW is the product name - for the convinience of the user, the program don't give !@#$ about it.
3. All the other rows from the THIRD is a client with the first cell being the name of the client, under each product you put in the quantity each client wants to buy.
4. If you have a blank cell the program fill in 0
