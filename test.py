
# Import date class from datetime module
from datetime import date
 
# Returns the current local date
date_obj = date.today()
date_string = date_obj.strftime("%B %d, %Y")
#date_time = today.strftime("%m/%d/%Y")
print(date_string)