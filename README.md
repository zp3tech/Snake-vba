# Snake-vba
Recreating the classic retro game of snake on an Excel spreadsheet using nothing but VBA. No tutorial was used; I created for more practice in designing something from scratch in a new language, application, and environment.

## INSTRUCTIONS TO SETUP & PLAY

### Manually setup board and controller
1. Create a macro-enabled Excel spreadsheet
1. Set width of columns A through AA to 2.57 (23 pixels)
1. Type "x" in cells A1:A27,B1:AA1,B27:AA27,AA2:AA26 to create the edges for the game board (Sub **BoardSetup()** will turn background black).
1. Type "UP" in cell I30, "RIGHT" in cell J31, "DOWN" in cell I32, and "LEFT" in cell H31
1. Then copy all the SnakeSubs.bas code from this repo into your Modules folder.
1. Create a button and assign it the macro **StartGame**

### Gameplay
Simply click the **StartGame** button then use your arrow keys on keyboard to control the snake and collect as many O's as you can without hitting the walls or your own body!
