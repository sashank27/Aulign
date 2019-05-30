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
    file = "test2.zmx"
    testFile = cwd + "\\" + file
    TheSystem.LoadFile(testFile, False)
    print("File Imported")

    # Get Surfaces
    TheLDE = TheSystem.LDE

    SurfaceList = []
    for i in range(TheLDE.NumberOfSurfaces):
        SurfaceList.append(TheLDE.GetSurfaceAt(i))

    TheLDE.GetSurfaceAt(9).TiltDecenterData.BeforeSurfaceOrder = constants.TiltDecenterOrderType_Decenter_Tilt
    # TheLDE.GetSurfaceAt(9).TiltDecenterData.BeforeSurfaceDecenterX = 2
    # TheLDE.GetSurfaceAt(9).TiltDecenterData.AfterSurfaceDecenterX = -2
    # TheLDE.GetSurfaceAt(9).TiltDecenterData.BeforeSurfaceDecenterY = -3
    # TheLDE.GetSurfaceAt(9).TiltDecenterData.AfterSurfaceDecenterY = 3

    cbData = TheLDE.GetSurfaceAt(11)
    print(cbData.Thickness)
    # print(cbData.MechanicalSemiDiameter)
    print(cbData.Parameter1)

    # print(help(TheLDE.GetSurfaceAt(9).TiltDecenterData))
    # print(TheLDE.GetSurfaceAt(8).TiltDecenterData.DecenterX)
    # print(TheLDE.GetSurfaceAt(9).TiltDecenterData.AfterSurfaceDecenterX)
    # print(TheLDE.GetSurfaceAt(9).TiltDecenterData.BeforeSurfaceDecenterY)
    # print(TheLDE.GetSurfaceAt(9).TiltDecenterData.AfterSurfaceDecenterY)

    CoordinateBreakSurfaces = [8, 11, 12, 15, 16, 19]
    for i in range(len(CoordinateBreakSurfaces), 2):
        Surface = SurfaceList[i]

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
    # print(dir(spot_results.SpotData))
    # print(help(spot_results.SpotData))
    print('RMS radius: %6.3f' %
          (spot_results.SpotData.GetRMSSpotSizeFor(1, 1)))
    print('GEO radius: %6.3f' %
          (spot_results.SpotData.GetGeoSpotSizeFor(1, 1)))

    # Set up Batch Ray Trace
    raytrace = TheSystem.Tools.OpenBatchRayTrace()
    nsur = TheSystem.LDE.NumberOfSurfaces
    max_rays = 30
    normUnPolData = raytrace.CreateNormUnpol(
        (max_rays + 1) * (max_rays + 1), constants.RaysType_Real, nsur)

    # Define batch ray trace constants
    hx = 0.0
    max_wave = TheSystem.SystemData.Wavelengths.NumberOfWavelengths
    num_fields = TheSystem.SystemData.Fields.NumberOfFields
    # print(max_wave, " ", num_fields)
    hy_ary = np.array([0, 0.707, 1])

    # Initialize x/y image plane arrays
    x_ary = np.empty((num_fields, max_wave, ((max_rays + 1) * (max_rays + 1))))
    y_ary = np.empty((num_fields, max_wave, ((max_rays + 1) * (max_rays + 1))))
    # print("x_ary shape", x_ary.shape)
    # print("x_ary shape", y_ary.shape)

    # Determine maximum field in Y only
    max_field = 0.0
    for i in range(1, num_fields + 1):
        if (TheSystem.SystemData.Fields.GetField(i).Y > max_field):
            max_field = TheSystem.SystemData.Fields.GetField(i).Y

    # print("Max field", max_field)
    plt.rcParams["figure.figsize"] = (5, 5)
    colors = ('k', 'g', 'r')
    markers = ('+', 's', '^')

    if TheSystem.SystemData.Fields.GetFieldType() == constants.FieldType_Angle:
        field_type = 'Angle'
    elif TheSystem.SystemData.Fields.GetFieldType() == constants.FieldType_ObjectHeight:
        field_type = 'Height'
    elif TheSystem.SystemData.Fields.GetFieldType() == constants.FieldType_ParaxialImageHeight:
        field_type = 'Height'
    elif TheSystem.SystemData.Fields.GetFieldType() == constants.FieldType_RealImageHeight:
        field_type = 'Height'

    # print("Field type", field_type)
    field = 1
    for wave in range(1, max_wave + 1):
        # Adding Rays to Batch, varying normalised object height hy
        normUnPolData.ClearData()
        waveNumber = wave
        # for i = 1:((max_rays + 1) * (max_rays + 1))
        for i in range(1, (max_rays + 1) * (max_rays + 1) + 1):

            px = np.random.random() * 2 - 1
            py = np.random.random() * 2 - 1

            while (px*px + py*py > 1):
                py = np.random.random() * 2 - 1
            normUnPolData.AddRay(
                waveNumber, hx, hy_ary[field - 1], px, py, constants.OPDMode_None)

        baseTool = CastTo(raytrace, 'ISystemTool')
        baseTool.RunAndWaitForCompletion()

        # Read batch raytrace and display results
        normUnPolData.StartReadingResults()
        output = normUnPolData.ReadNextResult()

        while output[0]:                                                    # success
            # ErrorCode & vignetteCode
            if ((output[2] == 0) and (output[3] == 0)):
                x_ary[field - 1, wave - 1, output[1] - 1] = output[4]   # X
                y_ary[field - 1, wave - 1, output[1] - 1] = output[5]   # Y
            output = normUnPolData.ReadNextResult()
        temp = plt.plot(np.squeeze(x_ary[field - 1, wave - 1, :]), np.squeeze(
            y_ary[field - 1, wave - 1, :]), '.', ms=3, c=colors[wave - 1],
            marker=markers[wave - 1])

    plt.title('Spot Diagram: %s' % (os.path.basename(testFile)))
    plt.draw()

    '''This will clean up the connection to OpticStudio.
    Note that it closes down the server instance of OpticStudio,
    so you for maximum performance do not do this until you need to.'''
    del zosapi
    zosapi = None

    # place plt.show() after clean up to release OpticStudio from memory
    # plt.show()
