# Excels
Some Excel projects with VBA

1st PROJECT
It is a Pathfinder 2nd Edition Alchemist template with automated functionalities.

2nd PROJECT 
It is a Monopoly Stock Exchange calculator. The board game has a Calculator on its own but I made this macro excel to replace it. It keeps track of the players' stocks and their value change over time around a fixed value. Additionally the stock values of a company increase or decrease if you buy or sell stocks from it respectively through a logic that I reverse engineered from the game's real calculator. It also displays the companies rents. Lastly there are some VBA macro buttons to automatically implement some specific Action Cards or transactions(rent, buy/sell stocks, how much you take if you pass the start etc) or reset the game.

Cells S to AR are hidden and contain the mechanism through which the stock values fluctuate over time and also a mechanism that solves a tie break if two players have equal stocks in order to determine the company's presidency/ownership (cannot be done with max function because in case of same value excel uses the left rule and in order to take ownwership of a company from someone you don't just need to buy the same stocks but to surpass that number, so if you reach the same stock number as the company's owner the other player continues to be the ownner regardless of whether his is left or right. This is done by keeping track of when you bought the stock and in such disputes the owner always bought the stocks less recently.)
