import glob
import os

path_list = r"O:\Field Services Division\Field Support Center\Project Acceptance\99999- PlaceHolder Until I do the sql\Pressurized-Pipe\*(Letter).pdf"
list_of_files = glob.glob(path_list) # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)
print(latest_file)

path_asset = r"O:\Field Services Division\Field Support Center\Project Acceptance\99999- PlaceHolder Until I do the sql\Pressurized-Pipe\*(Asset List).pdf"
list_of_files = glob.glob(path_asset) # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)
print(latest_file)