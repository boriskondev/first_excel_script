from datetime import datetime

time_start = datetime.now()
time_elapsed = datetime.now()
time_took = time_elapsed - time_start
print(f"The execution of this script {time_took.seconds} seconds.")