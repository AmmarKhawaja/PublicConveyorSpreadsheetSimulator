#can edit these variables
# - max number of times to run (in case of errors)
RUN_TIMES = 1000000
# - manual or random input
INPUT = "random"
# - input if manual input is selected
CONVEYOR_ADD = [.75,.50,.75,.50,.0,.25,.75,.75,
                .75,.75,.75,.0,.0,.75,.50,.75,
                .50,.25,.50,.25,.75,.75,.75,.25,]
# - probabilities if random input is selected [0, .25, .50, 75]
PROBS = [1, 2,]
# - number of inputs if random input is selected
INPUT_NUM = 100
# - number of sections in each zone
ZONE_LENGTH = 12
# - number of zones
ZONE_NUMBER = 8
# - turn algorithm on
ALGO_ON = True
# - should generate spreadsheet
SPREADSHEET = True
# - should program style spreadsheet (time)
STYLE = True
# - file to save to
FILE = "example run"

#cannot edit these variables
#Set up conveyor environment
CONVEYOR = []
EMPTY_CONVEYOR = []
ZONE_WEIGHTS = []
ZONE_RATING = []
SPEED_RATING = []
SPEED = []
SUB = " "