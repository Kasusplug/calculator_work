from whole_outputs import info_output, error_output, debug_output


class Logs:

    def __init__(self, file_name_read, file_name_record):
        self.file_name_read = file_name_read
        self.file_name_record = file_name_record
        self.logs_time_count_info = {}
        self.logs_time_count_error = {}
        self.logs_time_count_debug = {}
        self.count_info = {msg: 0 for msg in info_output}
        self.count_error = {msg: 0 for msg in error_output}
        self.count_debug = {msg: 0 for msg in debug_output}

    def logs_find(self):
        with open(self.file_name_read, 'r', encoding='utf-8') as logs_events:
            with open(self.file_name_record, 'w', encoding='utf-8') as new_logs:        
                for line in logs_events:
                    line = line.strip()
                    if "[INFO]" in line:
                        timestamp = line[0:90] 
                        if timestamp in self.logs_time_count_info:
                            self.logs_time_count_info[timestamp] += 1  
                        else:
                            self.logs_time_count_info[timestamp] = 1  
                        new_logs.write(line + '\n')
                    elif "[ERROR]" in line:
                        timestamp = line[0:90]
                        if timestamp in self.logs_time_count_error:
                            self.logs_time_count_error[timestamp] += 1  
                        else:
                            self.logs_time_count_error[timestamp] = 1  
                        new_logs.write(line + '\n')
                    elif "[DEBUG]" in line:
                        timestamp = line[0:90]
                        if timestamp in self.logs_time_count_debug:
                            self.logs_time_count_debug[timestamp] += 1  
                        else:
                            self.logs_time_count_debug[timestamp] = 1  
                        new_logs.write(line + '\n')
                print('Successfully processed all logs')
    
    def logs_counter(self):
        with open(r'task_1try\catched_logs.log', 'w', encoding='utf-8') as new_logs:
            for timestamp, count in self.logs_time_count_info.items():
                new_logs.write(f"  {timestamp} - {count}" + '\n')
            for timestamp, count in self.logs_time_count_error.items():
                new_logs.write(f"  {timestamp} - {count}" + '\n')
            for timestamp, count in self.logs_time_count_debug.items():
                new_logs.write(f"  {timestamp} - {count}" + '\n')
            print('Succesfully counted ALL logs')


    def logs_compile(self):
        with open(r'task_1try\catched_logs.log', 'r', encoding='utf-8') as read_logs:
            with open(r'task_1try\counted.log', 'w', encoding='utf-8') as counted_logs:
                for line in read_logs:
                    line = line.strip()

                    for info in info_output:
                        if info in line:
                            self.count_info[info] += 1
                    for error in error_output:
                        if error in line:
                            self.count_error[error] += 1
                    for debug in debug_output:
                        if debug in line:
                            self.count_debug[debug] += 1

                counted_logs.write("INFO logs count:\n")
                for info, count in self.count_info.items():
                    counted_logs.write(f"{info}: {count}\n")

                counted_logs.write("\nERROR logs count:\n")
                for error, count in self.count_error.items():
                    counted_logs.write(f"{error}: {count}\n")

                counted_logs.write("\nDEBUG logs count:\n")
                for debug, count in self.count_debug.items():
                    counted_logs.write(f"{debug}: {count}\n")

        print("whole outputs was counted successfully and written to file")


    def error_logs_write(self):
        with open(r'task_1try\catched_logs.log', 'r', encoding='utf-8') as read_logs:
            with open(r'task_1try\error.log', 'w', encoding='utf-8') as error_logs:
                for error_log in read_logs:
                    if '[ERROR]' in error_log:
                        error_logs.write(error_log)


    def counted_error_logs_write(self):
        with open(r'task_1try\counted.log', 'r', encoding='utf-8') as counted_logs:
            with open(r'task_1try\error_counted.log', 'w', encoding='utf-8') as error_counted_logs:
                for error_log in counted_logs:
                    if '[ERROR]' in error_log:
                        error_counted_logs.write(error_log)


check_logs = Logs(
    file_name_read=r'C:\Users\kasus\Desktop\pythonapps\logsbot\task_1try\logs.log',
    file_name_record=r'C:\Users\kasus\Desktop\pythonapps\logsbot\task_1try\catched_logs.log'
)

check_logs.logs_find()
check_logs.logs_counter()
check_logs.logs_compile()
check_logs.error_logs_write()
check_logs.counted_error_logs_write()

