# You are rolling 'm' number of dice, each with 'n' number of sides
# When the dice land, your dice value is whatever the largest die says
# What is the average roll's value?
#
# Most of the code below attempts to arrive at that number thru brute force.
# I then compare that brute force approach with matt parker's formula,
# which is:  avg = ((m / (m+1)) * n) + 0.5
# 
# SOURCE: https://www.youtube.com/watch?v=X_DdGRjtwAo&t=1484s

import random
import statistics

# adjust these values as you see fit
numSides = 6         # number of sides on each die
numDice  = 2         # number of dice for each roll
numRolls = 1000000   # number of rolls

# for each roll of the dice, use the highest value
highDice = [] 
for i in range(numRolls):
    thisRoll = []
    for ii in range(numDice):
        thisRoll.append(random.randint(1,numSides))
    highDice.append(max(thisRoll))
    
    # just a progress bar to show that things aren't stuck
    if i % (numRolls / 10) == 0:
        print(i)

# print the final result, the average highest die per roll
avg = mean(highDice) 
print('calculated average roll: ' + str(avg))

# print the results of matt parker's formula
mattSays = ((numDice / (numDice + 1)) * numSides) + 0.5
print('matt parker says it should be: ' + str(mattSays))
