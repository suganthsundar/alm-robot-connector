import win32com.client
import pythoncom


# ----------------------------------------------------
# Main Test Function
# Debug - Boolean. Equals to false if running in [Test Mode] : reporting to Application Lifecycle Management
# CurrentTestSet - [OTA COM Library].TestSet.
# CurrentTSTest - [OTA COM Library].TSTest.
# CurrentRun - [OTA COM Library].Run.
# ----------------------------------------------------
def Test_Main(Debug, CurrentTestSet, CurrentTSTest, CurrentRun):
  try :
    # clear output window
    TDOutput.Clear()

    # TODO: put your code here
    import sys, os
    WORKSPACE = r'C:\Comcast_QA'
    sys.path.insert(0, os.path.join(WORKSPACE, 'spares-test'))

    ALM_OBJECTS = {
        "Debug": Debug,
        "CurrentTestSet": CurrentTestSet,
        "CurrentTSTest": CurrentTSTest,
        "CurrentRun": CurrentRun,
        "TDHelper": TDHelper,
        "TDInput":  TDInput,
        "TDOutput": TDOutput,
        "XTools": XTools
    }

    import runner
    result = runner.run(r'tests\api\advance_search\network_location.txt', **ALM_OBJECTS)

    if not Debug :
      TDOutput.Print("Put your code here")
  # handle errors
  except pythoncom.com_error, (hr, desc, exc, arg):
    TDOutput.Print("Run-time error %d: %s" % (hr, desc))
    # update execution status in "Test" mode
    if not Debug :
      CurrentRun.Status = "Failed"
      CurrentTSTest.Status = "Failed"
