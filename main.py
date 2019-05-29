from win32com.client.gencache import EnsureDispatch, EnsureModule
from win32com.client import CastTo, constants
import os
import matplotlib.pyplot as plt
import numpy as np
import inspect


class PythonStandaloneApplication(object):

    class LicenseException(Exception):
        pass

    class ConnectionException(Exception):
        pass

    class InitializationException(Exception):
        pass

    class SystemNotPresentException(Exception):
        pass

    def __init__(self):
        '''
        make sure the Python wrappers are available for the COM client and
        interfaces
        # EnsureModule('ZOSAPI_Interfaces', 0, 1, 0)
        Note - the above can also be accomplished using 'makepy.py' in the
        following directory:
            {PythonEnv}\Lib\site-packages\wind32com\client\
        Also note that the generate wrappers do not get refreshed when the
        COM library changes.
        To refresh the wrappers, you can manually delete everything in the
        cache directory:
           {PythonEnv}\Lib\site-packages\win32com\gen_py\*.*
        '''

        self.TheConnection = EnsureDispatch("ZOSAPI.ZOSAPI_Connection")
        if self.TheConnection is None:
            raise PythonStandaloneApplication.ConnectionException(
                "Unable to intialize COM connection to ZOSAPI")

        self.TheApplication = self.TheConnection.CreateNewApplication()
        if self.TheApplication is None:
            raise PythonStandaloneApplication.InitializationException(
                "Unable to acquire ZOSAPI application")

        if self.TheApplication.IsValidLicenseForAPI is False:
            raise PythonStandaloneApplication.LicenseException(
                "License is not valid for ZOSAPI use")

        self.TheSystem = self.TheApplication.PrimarySystem
        if self.TheSystem is None:
            raise PythonStandaloneApplication.SystemNotPresentException(
                "Unable to acquire Primary system")

    def __del__(self):
        if self.TheApplication is not None:
            self.TheApplication.CloseApplication()
            self.TheApplication = None

        self.TheConnection = None

    def OpenFile(self, filepath, saveIfNeeded):
        if self.TheSystem is None:
            raise PythonStandaloneApplication.SystemNotPresentException(
                "Unable to acquire Primary system")
        self.TheSystem.LoadFile(filepath, saveIfNeeded)

    def CloseFile(self, save):
        if self.TheSystem is None:
            raise PythonStandaloneApplication.SystemNotPresentException(
                "Unable to acquire Primary system")
        self.TheSystem.Close(save)

    def SamplesDir(self):
        if self.TheApplication is None:
            raise PythonStandaloneApplication.InitializationException(
                "Unable to acquire ZOSAPI application")

        return self.TheApplication.SamplesDir

    def ExampleConstants(self):
        if self.TheApplication.LicenseStatus is constants.LicenseStatusType_PremiumEdition:
            return "Premium"
        elif self.TheApplication.LicenseStatus is constants.LicenseStatusType_ProfessionalEdition:
            return "Professional"
        elif self.TheApplication.LicenseStatus is constants.LicenseStatusType_StandardEdition:
            return "Standard"
        else:
            return "Invalid"


if __name__ == '__main__':
    zosapi = PythonStandaloneApplication()
    value = zosapi.ExampleConstants()

    TheSystem = zosapi.TheSystem
    TheApplication = zosapi.TheApplication

    # Set up primary optical system
    cwd = os.getcwd()
    file = "test.zmx"
    testFile = cwd + "\\" + file
    TheSystem.LoadFile(testFile, False)
    print("File Imported")

    # Get Surfaces
    TheLDE = TheSystem.LDE

    SurfaceList = []
    for i in range(TheLDE.NumberOfSurfaces):
        SurfaceList.append(TheLDE.GetSurfaceAt(i))

    # Spot Diagram Analysis Results
    spot = TheSystem.Analyses.New_Analysis(constants.AnalysisIDM_StandardSpot)
    spot_setting = spot.GetSettings()
    baseSetting = CastTo(spot_setting, 'IAS_Spot')
    baseSetting.Field.SetFieldNumber(0)
    baseSetting.Wavelength.UseAllWavelengths()
    baseSetting.Surface.UseImageSurface()

    # extract RMS & Geo spot size for field points
    base = CastTo(spot, 'IA_')
    base.ApplyAndWaitForCompletion()
    spot_results = base.GetResults()
    print(dir(spot_results.SpotData))
    print(help(spot_results.SpotData))
    print('RMS radius: %6.3f' %
          (spot_results.SpotData.GetRMSSpotSizeFor(1, 1)))
    print('GEO radius: %6.3f' %
          (spot_results.SpotData.GetGeoSpotSizeFor(1, 1)))

    # analIDM = []
    # API_enum = list(constants.__dicts__[0].keys())
    # for i in API_enum:
    #     if i.find('AnalysisIDM_') != -1:
    #         analIDM.append(constants.__dicts__[0].get(i))
    #         print(constants.__dicts__[0].get(i), ':  ', i)
    # analIDM.sort()

    # for k in analIDM:
    #     a = TheSystem.Analyses.New_Analysis(k)
    #     if a is None:
    #         print('This analysis cannot be opened in ',
    #               'Sequential Mode' if TheSystem.Mode == 0 else 'Non-Sequential Mode',
    #               ': enumID ', k)
    #         continue
    #     print(a.GetAnalysisName, '\t', a.HasAnalysisSpecificSettings)
    #     a.Close()

    '''This will clean up the connection to OpticStudio.
    Note that it closes down the server instance of OpticStudio,
    so you for maximum performance do not do this until you need to.'''
    del zosapi
    zosapi = None
