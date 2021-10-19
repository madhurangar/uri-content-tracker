import openpyxl, sys
from tqdm.auto import tqdm, trange
from utils import utils as u

INPUT =  'data.xlsx'

wb = openpyxl.load_workbook(INPUT, data_only=True)
ws = wb['Sheet1']

# loop over all rows except the header
for row in trange(2, ws.max_row+1):
    try:
        u.execute_row(ws, row)
    except KeyboardInterrupt:
        # tqdm.write("Exiting...")
        sys.exit()

# save file and exit
print("Saving...")
wb.save(INPUT)

print("Done!")
