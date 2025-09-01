import pyautogui
import ctypes
import time
import pytesseract
import pandas as pd
from pynput.keyboard import Controller

pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract'

class text:
   BOLD = '\033[1m'
   UNDERLINE = '\033[4m'
   END = '\033[0m'

def switch_to_application(app_title):
    """
    switch_to_application: switches to the application with the nearest name in question
    """
    try:
        # Get all open windows
        windows = pyautogui.getAllWindows()
        # Find the window with the matching title
        target_window = None
        for window in windows:
            if app_title.lower() in window.title.lower(): # Case-insensitive search
                target_window = window
                break

        if target_window:
            hWnd = target_window._hWnd
            user32 = ctypes.windll.user32

            # If minimized, restore/display window
            if user32.IsIconic(hWnd):
                user32.ShowWindow(hWnd, 5) # SW_RESTORE

            # Activates the Window
            user32.SetForegroundWindow(hWnd)
            print(f"Switched to '{target_window.title}'")
        else:
            print(f"Application with title '{app_title}' not found.")

    except Exception as e:
        print(f"An error occurred: {e}")
def check_ingredients(region_val, ingredients, effects):
    """check_ingredients: records the current known info/effects of ingredients"""
    current_ingredient_info = pd.DataFrame(columns=["Ingredient", "Effect 1", "Effect 2", "Effect 3", "Effect 4"])
    
    # loop over every ingredient
    for i in range(188):
        image = pyautogui.screenshot(region=region_val)
        txt = pytesseract.image_to_string(image)
        txt = txt.replace("’", "'")
        found_ingredient = [ingredient for ingredient in ingredients if ingredient.lower() in txt.lower()]
        if'Blue Butterfly Wing' in found_ingredient:
            found_ingredient.remove('Butterfly Wing')
        # exits if the same ingredient is found meaning you have reached the bottom of the list
        if(current_ingredient_info["Ingredient"].eq(found_ingredient[0]).any()):
            break
        print(f"Ingredient Found: {found_ingredient}")
        found_effects = [effect for effect in effects if effect.lower() in txt.lower()]
        if 'Damage Magicka Regen' in found_effects or 'Lingering Damage Magicka' in found_effects:
            found_effects.remove('Damage Magicka')
        if 'Damage Stamina Regen' in found_effects or 'Lingering Damage Stamina' in found_effects:
            found_effects.remove('Damage Stamina')
        if 'Lingering Damage Health' in found_effects:
            found_effects.remove('Damage Health')
        while len(found_effects) < 4:
            found_effects.append('Unknown')
        print(f"Effect(s) Found: {found_effects}")
        new_ingredient = pd.DataFrame({
            "Ingredient": found_ingredient,
            "Effect 1": found_effects[0],
            "Effect 2": found_effects[1],
            "Effect 3": found_effects[2],
            "Effect 4": found_effects[3]
        })
        current_ingredient_info = pd.concat([current_ingredient_info, new_ingredient], ignore_index=True)
        keyboard.press('s')
        keyboard.release('s')
        time.sleep(.05) # wait for the key press before taking another screenshot
    return current_ingredient_info
def define_ingredient_dictionary(region_val, ingredients, effects):
    """define_ingredeint_dictionary: exextutes switching to the application and checking ingredients process.
        this function can be used to 'refresh' the ingredient list
    """
    top = False
    previous_ingredient = None
    switch_to_application("Skyrim")
    time.sleep(3) # wait for application to open
    while not top: # takes you to the top of the application
        keyboard.press('w')
        image = pyautogui.screenshot(region=region_val)
        txt = pytesseract.image_to_string(image)
        txt = txt.replace("’", "'") 
        found_ingredient = [ingredient for ingredient in ingredients if ingredient.lower() in txt.lower()]
        if 'Blue Butterfly Wing' in found_ingredient:
            found_ingredient.remove('Butterfly Wing')
        if found_ingredient[0] == previous_ingredient:
            keyboard.release('w')
            top = True
        previous_ingredient = found_ingredient[0]
        time.sleep(.05)
    return check_ingredients(region_val, ingredients, effects)
def find_effect_matches(known_ingredients, target):
    """find_effect_matches: finds ingredients with the same effects as the target ingredient"""
    columns_to_check = ['Effect 1', 'Effect 2', 'Effect 3', 'Effect 4']
    ingredient_list = []
    for i in range(1, 5):
        row = known_ingredients.loc[known_ingredients['Ingredient'] == target]
        if not row.size:
            print("Invalid ingredient entered. Make sure the spelling is correct with apostrophes!")
            return
        effect = row.iloc[0][f'Effect {i}']
        if effect == 'Unknown':
            break
        rows_with_value = (known_ingredients[columns_to_check] == effect).any(axis=1)
        result = known_ingredients[rows_with_value] # removes target from list
        result = result[result['Ingredient'] != target]
        if result.size: 
            ingredient_list.extend(result['Ingredient'].to_list())
    ingredient_list = list(set(ingredient_list)) # removes duplicates
    ingredient_list.sort() # alphabetizes
    [print(ingredient) for ingredient in ingredient_list]

keyboard = Controller()
# sets dimensions of screenshots
screen_width, screen_height = pyautogui.size()
width = round(screen_width * 0.5) # finds pixel percentage of screen width
height = round(screen_height * 0.6)
top_left_x = screen_width - width
top_left_y = screen_height - height - 200
region_val = (top_left_x,top_left_y,width,height)
# finds the names of all ingredients and effects from website
info_url = "https://elderscrolls.fandom.com/wiki/Ingredients_(Skyrim)"
all_tables = pd.read_html(info_url)
ingredients = pd.concat([all_tables[1], all_tables[2], all_tables[3], all_tables[4], all_tables[5]])['Ingredient'].tolist()
effects = all_tables[6]['Potion'].tolist() 
known_ingredients = define_ingredient_dictionary(region_val, ingredients, effects)
target_ingredient = None
print("\n")
while True:
    print("-------------------------------------------------------------------------")
    print("Type 'Refresh' to Refresh the Ingredient Matcher. Type 'Quit' to Quit.")
    target_ingredient = input("Target Ingredient: ")
    target_ingredient = target_ingredient.title()
    print()
    match target_ingredient:
        case 'Refresh':
            known_ingredients = define_ingredient_dictionary(region_val, ingredients, effects)
        case 'Quit':
            break
        case _:
            find_effect_matches(known_ingredients, target_ingredient)
    print()



