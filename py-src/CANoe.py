# -----------------------------------------------------------------------------
# CANoe COM object
# -----------------------------------------------------------------------------

from win32com.client import *
from win32com.client.connect import *
from queue import Queue
        
class Canoe(object):
    """class for CANoe Application object, sync to MATLAB"""
    
    def __init__(self):
        self.App = None

    def get_application(self, visible=True):
        # link to exist CANoe
        self.App = Dispatch('CANoe.Application')
        self.App.Visible = visible
    
    def get_SysVar(self, ns_name, sysvar_name):

        if (self.App != None):
            systemNamespaces = self.App.System.Namespaces
            sys_namespace = systemNamespaces(ns_name)
            sys_value = sys_namespace.Variables(sysvar_name)
            return sys_value.Value
        else:
            raise RuntimeError("CANoe is not open,unable to GetVariable")

    def set_SysVar(self, ns_name, sysvar_name, val):

        if (self.App != None):
            systemNamespaces = self.App.System.Namespaces
            sys_namespace = systemNamespaces(ns_name)
            sys_value = sys_namespace.Variables(sysvar_name)
            sys_value.Value = val
        else:
            raise RuntimeError("CANoe is not open,unable to GetVariable")

    def get_test_result(self, test_config_name='Test_Configuration_1'):
        return self.get_SysVar(test_config_name,'VerdictSummary')

    def is_running(self):
        return self.App.Measurement.Running
    
    def open_canoe_config(self, cfg_path):
        if len(cfg_path) < 4 or cfg_path[-4:]!='.cfg':
            raise RuntimeError('Not correct CANoe cfg file')
        cfg_name = cfg_path.split('/')[-1]
        if self.App.Configuration.Name != cfg_name[:-4]:
            self.App.Open(cfg_path)

    def load_test_env(self, test_path, power_on):
        if len(test_path)<4 or test_path[-4:] != '.tse':
            raise RuntimeError('Not correct Test Environment file')
        
        if self.App.Configuration.TestSetup.TestEnvironments.Count == 1:
            pass
        else:
            while self.App.Configuration.TestSetup.TestEnvironments.Count != 0:
                self.App.Configuration.TestSetup.TestEnvironments.Remove(1,False)
            self.App.Configuration.TestSetup.TestEnvironments.Add(test_path)
        self._test_env_check()
        tmseq = self.test_module.Sequence
        tg = CastTo(tmseq(1),"ITestGroup")
        tgseq = tg.Sequence
        if power_on == 'On':
            tgseq('PowerOn').Enabled = True
            tgseq('PowerOff').Enabled = False
        else:
            tgseq('PowerOn').Enabled = False
            tgseq('PowerOff').Enabled = True
        
    def _test_env_check(self):
        self.tes = self.App.Configuration.TestSetup.TestEnvironments
        if self.tes.Count != 1:
            raise RuntimeError("Not a Test Configuration.")
        self.test_env = self.tes(1)
        self.test_env = CastTo(self.test_env, "ITestEnvironment2")
        test_modules = self.test_env.TestModules
        self.test_module = test_modules(1)
        self.test_m_report = self.test_module.Report
        self.test_m_report.Enabled = False
        
        
    def run_test_module(self):
        self.test_module.Start()
        
    def start_meas(self):
        self.App.Measurement.Start()
        
    def stop_meas(self):
        self.App.Measurement.Stop()
    
    def remove_test_env(self):
        self._test_env_check()
        self.App.Configuration.TestSetup.TestEnvironments.Remove(1,False)
        
    def load_test_config(self, test_path):
        if len(test_path)<7 or test_path[-7:] != '.vtuexe':
            raise RuntimeError('Not correct Test Unit file')
        if self.App.Configuration.TestConfigurations.Count != 0:
            while self.App.Configuration.TestConfigurations.Count != 0:
                self.App.Configuration.TestConfigurations.Remove(1)

        test_config = self.App.Configuration.TestConfigurations.Add()
        test_units = CastTo(test_config.TestUnits, "ITestUnits2")
        test_units.Add(test_path)

    def _test_config_check(self):
        tcs = self.App.Configuration.TestConfigurations
        if tcs.Count != 1:
            raise RuntimeError("Not a Test Configuration.")
        self.test_config = CastTo(tcs(1), "ITestConfiguration7")
        self.test_config_settings = CastTo(self.test_config.Settings, "ITestConfigurationSettings4")
        self.test_config_report = CastTo(self.test_config.Report, "ITestConfigurationReport4")
        
        if self.test_config.TestUnits.Count != 1:
            raise RuntimeError("Not a Test Unit")
        self.test_units = CastTo(self.test_config.TestUnits, "ITestUnits2")
        self.test_unit = CastTo(self.test_units(1), "ITestUnit3")
        self.test_unit_report = CastTo(self.test_unit.Report, "ITestUnitReport3")        
            
    def enable_test_case(self, tc_name):
        """Load exist test configuration, and set specified test case enabled."""
        self._set_time = 0
        self._test_config_check()

        self.test_case = tc_name
        gen = self._traverse_test_unit(self.test_unit, self._check_test_case_name)
        for _ in gen:
            pass

        if self._set_time > 1:
            raise RuntimeError("More than one Test Case is enabled.")
        elif self._set_time < 1:
            raise RuntimeError("No Test Case is enabled.")
            
    def _traverse_test_unit(self, test_unit, func=None):
        if test_unit.Elements.Count == 0:
            raise RuntimeError("No elements in Test Unit")
        result_queue = Queue()
        
        for elements in test_unit.Elements:
            self._find_subelement(result_queue, elements)
        

        while not result_queue.empty():
            yield func(result_queue.get())            
            
    def _find_subelement(self, queue, element):
        if any(element.Elements):
            for te in element.Elements:
                self._find_subelement(queue, te)
        else:
            queue.put(element)
            
    def _get_test_case_name(self, test_element):
        return test_element.Caption

    def _check_test_case_name(self, test_element):
        if test_element.Caption == self.test_case:
            test_element.Enabled = True
            self._set_time += 1
        else:
            test_element.Enabled = False


    def append_symbol_mappings(self, mapping_file):
        
        mp_obj = self.App.Configuration.SymbolMappings
        mp_obj.Delete()
        _, wrong_message = mp_obj.Append(mapping_file)
        if wrong_message != '':
            raise RuntimeError("Error Mapping: wrong_message")
        
    def close_canoe(self):
        self.App.Quit()

    def save_canoe(self):
        self.App.Configuration.Save()

    def get_all_test_cases(self):
        """ Used to get all test cases name """
        self._test_config_check()
        return self._traverse_test_unit(self.test_unit, self._get_test_case_name)

    def set_test_config(self, var, options='default'):
        """ Need to be refactoring """
        self._test_config_check()

        if options == 'default':
            #self.test_config_settings.StartOnMeasurement = True
            self.test_config_settings.StartOnSysVar = var
            self.test_config_settings.IgnoreBreakOnFail = True
            self.test_config_report.UseJointReport = False

    def set_test_report(self, path):
        self._test_config_check()
        self.test_unit_report.FullPath = path
        self.test_unit_report.Enabled = True

    def import_variant_profile(self, variant_file):

        self.test_config.ImportVariantProfilesAsync(variant_file)


    def set_logging(self, logging_name):
        if self.App.Configuration.OnlineSetup.LoggingCollection.Count != 1:
            raise RuntimeError('Not a Logging block')
        self.App.Configuration.OnlineSetup.LoggingCollection(1).FullName = logging_name
        
