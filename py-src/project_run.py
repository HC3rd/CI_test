from CANoe import Canoe
import time, os


cur_wd, _ = os.path.split(os.path.realpath('__file__'))
print(cur_wd)
wd_list = cur_wd.split('\\')
par_wd = '\\'.join(wd_list[:-1])

os.chdir(par_wd)
print(os.getcwd())

canoe = Canoe()

canoe.get_application()
prj = '\\canoe-prj\\singleCAN.cfg'

canoe.open_canoe_config(par_wd+prj)
canoe.start_meas()
t_st = time.time()
time.sleep(4)
canoe.stop_meas()
t_ed = time.time()
print(f'time elapsed: {t_ed-t_st}')
canoe.save_canoe()
canoe.close_canoe()