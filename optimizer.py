from pulp import *
import pandas
from openpyxl import load_workbook
from datetime import datetime

# Aux functions
startTime = datetime.now()


def addMacroConstraint(prob, kcal, mealMacros, mealVars, meals, macroPerc, macroTolerance, isFat):
    prob += lpSum([mealMacros[f]*mealVars[f]*(9 if isFat else 4) for f in meals]) / \
        kcal - macroPerc >= -macroTolerance * macroPerc
    prob += lpSum([mealMacros[f]*mealVars[f]*(9 if isFat else 4) for f in meals]) / \
        kcal - macroPerc <= macroTolerance * macroPerc
    return prob


def addMealConstraint(prob, kcal, mealNames, protein, carbs, fat, mealVars, mealPerc, mealTolerance):
    prob += lpSum([(protein[f]*4 + carbs[f]*4 +
                    fat[f]*9)*mealVars[f] for f in {k: mealVars[k] for k in mealNames}])/kcal - mealPerc >= -(0 if mealPerc == 0 else mealTolerance) * mealPerc
    prob += lpSum([(protein[f]*4 + carbs[f]*4 +
                    fat[f]*9)*mealVars[f] for f in {k: mealVars[k] for k in mealNames}])/kcal - mealPerc <= (0 if mealPerc == 0 else mealTolerance) * mealPerc
    return prob


def addMealBalance(prob, mealName, mealVars, index1, index2, index3, index4, portionTolerance):
    prob += mealVars[mealName + '_Green'] >= (-portionTolerance * mealBalances[index1] + mealBalances[index1]) * (
        mealVars[mealName + '_Green'] + mealVars[mealName + '_Main'] + mealVars[mealName + '_Side'] + mealVars[mealName + '_Other'])
    prob += mealVars[mealName + '_Green'] <= (portionTolerance * mealBalances[index1] + mealBalances[index1]) * (
        mealVars[mealName + '_Green'] + mealVars[mealName + '_Main'] + mealVars[mealName + '_Side'] + mealVars[mealName + '_Other'])

    prob += mealVars[mealName + '_Main'] >= (-portionTolerance * mealBalances[index2] + mealBalances[index2]) * (
        mealVars[mealName + '_Green'] + mealVars[mealName + '_Main'] + mealVars[mealName + '_Side'] + mealVars[mealName + '_Other'])
    prob += mealVars[mealName + '_Main'] <= (portionTolerance * mealBalances[index2] + mealBalances[index2]) * (
        mealVars[mealName + '_Green'] + mealVars[mealName + '_Main'] + mealVars[mealName + '_Side'] + mealVars[mealName + '_Other'])

    prob += mealVars[mealName + '_Side'] >= (-portionTolerance * mealBalances[index3] + mealBalances[index3]) * (
        mealVars[mealName + '_Green'] + mealVars[mealName + '_Main'] + mealVars[mealName + '_Side'] + mealVars[mealName + '_Other'])
    prob += mealVars[mealName + '_Side'] <= (portionTolerance * mealBalances[index3] + mealBalances[index3]) * (
        mealVars[mealName + '_Green'] + mealVars[mealName + '_Main'] + mealVars[mealName + '_Side'] + mealVars[mealName + '_Other'])

    prob += mealVars[mealName + '_Other'] >= (-portionTolerance * mealBalances[index4] + mealBalances[index4]) * (
        mealVars[mealName + '_Green'] + mealVars[mealName + '_Main'] + mealVars[mealName + '_Side'] + mealVars[mealName + '_Other'])
    prob += mealVars[mealName + '_Other'] <= (portionTolerance * mealBalances[index4] + mealBalances[index4]) * (
        mealVars[mealName + '_Green'] + mealVars[mealName + '_Main'] + mealVars[mealName + '_Side'] + mealVars[mealName + '_Other'])
    return prob


# Define current person number
person = 0

# Define base tolerances
macroTolerance = 0.05  # df['macroTolerance'][0]
macroToleranceCeiling1 = 0.15  # df['portionTolerance'][0]
macroToleranceCeiling2 = 0.5  # df['portionTolerance'][0]
macroToleranceCeiling3 = 0.5  # df['portionTolerance'][0]
mealTolerance = 0.05  # df['mealTolerance'][0]
mealToleranceCeiling1 = 0.25  # df['mealTolerance'][0]
mealToleranceCeiling2 = 0.5  # df['portionTolerance'][0]
mealToleranceCeiling3 = 0.65  # df['portionTolerance'][0]
portionTolerance = 0.05  # df['portionTolerance'][0]
portionToleranceCeiling1 = 0.25  # df['portionTolerance'][0]
portionToleranceCeiling2 = 0.5  # df['portionTolerance'][0]
portionToleranceCeiling3 = 0.65  # df['portionTolerance'][0]

# Init results object
results = {'Date': [], 'Program': [], 'Meals': [], 'MealValues': [], 'Person': [],
           'macroTolerance': [], 'mealTolerance': [], 'portionTolerance': []}
# Init failed results object
failedResults = {'Date': [], 'Program': [], 'Person': []}

while True:
    # Select data for one person
    df = pandas.read_excel('data.xlsx', sheet_name='Input_sheet', nrows=21, skiprows=[] if person == 0 else [
        i for i in range(1, person*21+1)])

    try:
        personName = df['Person'][0]
        print('Person Name: ' + personName)
        date = df['Date'][0]
        print('Date: ' + str(date))
        program = df['Program'][0]
        print('Program: ' + str(program))
    except:
        print('Last meal calculated. Goodbye!')
        break

    # Create a list of the meal constituents
    meals = list(df['Meals'])

    # Create a dictionary of costs for all meal constituents
    costs = dict(zip(meals, df['U']))

    # Create a dictionary of each macro for all meal constituents
    protein = dict(zip(meals, df['P']))
    fat = dict(zip(meals, df['F']))
    carbs = dict(zip(meals, df['C']))

    # Get kcal target and macro distribution
    kcal = df['kcal'][0]
    print('kcal target ' + str(kcal))
    pPerc = df['pPerc'][0]
    cPerc = df['cPerc'][0]
    fPerc = df['fPerc'][0]

    # Get meal distribution

    afternoonSnack = df['Breakfast'][1]
    booster1 = df['Breakfast'][4]
    booster2 = df['Booster2'][7]
    breakfast = df['Breakfast'][10]
    dinner = df['Breakfast'][13]
    lunch = df['Breakfast'][16]
    morningSnack = df['Breakfast'][19]

    # Get meal balances
    mealBalances = df['Meal Balance']

    # Create meal variables
    mealVars = LpVariable.dicts(
        'Meal', meals, lowBound=0, cat='Continuous')

    # Create problem and main opjective function
    prob = LpProblem('Portion_Distribution_Model', LpMinimize)
    prob += lpSum([costs[i]*mealVars[i] for i in meals])

    print('Macro tolerance=' + str(macroTolerance))
    print('Meal tolerance=' + str(mealTolerance))
    print('Portion tolerance=' + str(portionTolerance))

    # MACROS
    # # kcal
    prob += lpSum([(protein[f] * 4 + carbs[f] * 4 + fat[f] * 9)
                   * mealVars[f] for f in meals]) == kcal
    # # protein
    prob = addMacroConstraint(
        prob, kcal, protein, mealVars, meals, pPerc, macroTolerance, False)
    # # carbs
    prob = addMacroConstraint(prob, kcal, carbs, mealVars,
                              meals, cPerc, macroTolerance, False)
    # # fat
    prob = addMacroConstraint(prob, kcal, fat, mealVars,
                              meals, fPerc, macroTolerance, True)

    # MEALS
    # afternoon snack
    prob = addMealConstraint(prob, kcal, (
        'Afternoon_Snack_Main', 'Afternoon_Snack_Green', 'Afternoon_Snack_Side', 'Afternoon_Snack_Other'), protein, carbs, fat, mealVars, afternoonSnack, mealTolerance)
    # # booster 1
    prob = addMealConstraint(prob, kcal, (
        'Booster1_Main', 'Booster1_Green', 'Booster1_Side', 'Booster1_Other'), protein, carbs, fat, mealVars, booster1, mealTolerance)
    # # booster 2
    prob = addMealConstraint(prob, kcal, (
        'Booster2_Main', 'Booster2_Green', 'Booster2_Side', 'Booster2_Other'), protein, carbs, fat, mealVars, booster2, mealTolerance)
    # # breakfast
    prob = addMealConstraint(prob, kcal, (
        'Breakfast_Main', 'Breakfast_Green', 'Breakfast_Side', 'Breakfast_Other'), protein, carbs, fat, mealVars, breakfast, mealTolerance)
    # # dinner
    prob = addMealConstraint(prob, kcal, (
        'Dinner_Main', 'Dinner_Green', 'Dinner_Side', 'Dinner_Other'), protein, carbs, fat, mealVars, dinner, mealTolerance)
    # # lunch
    prob = addMealConstraint(prob, kcal, (
        'Lunch_Main', 'Lunch_Green', 'Lunch_Side', 'Lunch_Other'), protein, carbs, fat, mealVars, lunch, mealTolerance)
    # # morning snack
    prob = addMealConstraint(prob, kcal, (
        'Morning_Snack_Main', 'Morning_Snack_Green', 'Morning_Snack_Side', 'Morning_Snack_Other'), protein, carbs, fat, mealVars, morningSnack, mealTolerance)

    # # MEAL BALANCE
    # afternoon snack
    prob = addMealBalance(prob, 'Afternoon_Snack', mealVars,
                          0, 1, 2, portionTolerance)
    # # booster 1
    prob = addMealBalance(prob, 'Booster1', mealVars,
                          3, 4, 5, portionTolerance)
    # # booster 2
    prob = addMealBalance(prob, 'Booster2', mealVars,
                          6, 7, 8, portionTolerance)
    # # breakfast
    prob = addMealBalance(prob, 'Breakfast', mealVars,
                          9, 10, 11, portionTolerance)
    # # dinner
    prob = addMealBalance(prob, 'Dinner', mealVars,
                          12, 13, 14, portionTolerance)
    # # lunch
    prob = addMealBalance(prob, 'Lunch', mealVars, 15,
                          16, 17, portionTolerance)
    # # morning snack
    prob = addMealBalance(prob, 'Morning_Snack', mealVars,
                          18, 19, 20, portionTolerance)
    prob.solve()
    status = LpStatus[prob.status]
    print('Status:', status)

    resultDict = {v.name.replace(
        'Meal_', ''): v.varValue for v in prob.variables()}
    kcalTotal = sum([(protein[f] * 4 + carbs[f] * 4 + fat[f] * 9)
                     * resultDict[f] for f in meals])
    print('Total KCAL = ' + str(kcalTotal))

    if status == 'Optimal':
        for v in prob.variables():
            results['Date'].append(str(date))
            results['Program'].append(program)
            results['Meals'].append(v.name)
            results['MealValues'].append(v.varValue)
            results['Person'].append(personName)
            results['macroTolerance'].append(macroTolerance)
            results['mealTolerance'].append(mealTolerance)
            results['portionTolerance'].append(portionTolerance)
        # Reset base tolerances
        macroTolerance = 0.05
        mealTolerance = 0.05
        portionTolerance = 0.05
        print('###########################')
        print('#########success###########')
        print('###########################')
        person += 10000
        print(prob)
    else:
        print('Failed with these tolerances, trying again...')
        person -= 1
        if portionTolerance < portionToleranceCeiling1:
            portionTolerance = round(portionTolerance + 0.05, 3)
        elif mealTolerance < mealToleranceCeiling1:
            mealTolerance = round(mealTolerance + 0.05, 3)
        elif macroTolerance < macroToleranceCeiling1:
            macroTolerance = round(macroTolerance + 0.05, 3)
        elif portionTolerance < portionToleranceCeiling2:
            portionTolerance = round(portionTolerance + 0.05, 3)
        elif mealTolerance < mealToleranceCeiling2:
            mealTolerance = round(mealTolerance + 0.05, 3)
        elif macroTolerance < macroToleranceCeiling2:
            macroTolerance = round(macroTolerance + 0.05, 3)
        elif portionTolerance < portionToleranceCeiling3:
            portionTolerance = round(portionTolerance + 0.05, 3)
        elif mealTolerance < mealToleranceCeiling3:
            mealTolerance = round(mealTolerance + 0.05, 3)
        elif macroTolerance < macroToleranceCeiling3:
            macroTolerance = round(macroTolerance + 0.05, 3)
        else:
            failedResults['Date'].append(str(date))
            failedResults['Program'].append(program)
            failedResults['Person'].append(personName)
            person += 1
            print('###########################')
            print('##########failed###########')
            print('###########################')

    person += 1


odf = pandas.DataFrame(results, columns=[
    'Date', 'Program', 'Meals', 'MealValues', 'Person', 'macroTolerance', 'mealTolerance', 'portionTolerance'])

frdf = pandas.DataFrame(failedResults, columns=['Date', 'Program', 'Person'])

book = load_workbook('data.xlsx')
# https://github.com/PyCQA/pylint/issues/3060
writer = pandas.ExcelWriter(  # pylint: disable=abstract-class-instantiated
    'data.xlsx')
writer.book = book
odf.to_excel(writer, sheet_name='Output_sheet')
frdf.to_excel(writer, sheet_name='Failed_sheet')
writer.save()
writer.close()


print('Run time: ' + str(datetime.now() - startTime))
