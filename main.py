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
        Surface = TheLDE.GetSurfaceAt(i)
        SurfaceList.append(Surface)
        # print(i)
        # print(Surface.Type)
        # print(Surface.TypeName == 'Coordinate Break')

    CoordinateBreakSurfaces = [8, 11, 12, 15, 16, 19]
    for i in range(len(CoordinateBreakSurfaces)):
        Surface = TheLDE.GetSurfaceAt(CoordinateBreakSurfaces[i])
        # SurfaceCast = CastTo(Surface.SurfaceData, 'ISurfaceCoordinateBreak')
        print(CoordinateBreakSurfaces[i])
        print((Surface.GetSurfaceCell(constants.SurfaceColumn_Par1)))
        print((Surface.GetSurfaceCell(constants.SurfaceColumn_Par2)))
        print((Surface.GetSurfaceCell(constants.SurfaceColumn_Par3)))
        print((Surface.GetSurfaceCell(constants.SurfaceColumn_Par4)))
        print((Surface.GetSurfaceCell(constants.SurfaceColumn_Par5)))

    TheLDE.GetSurfaceAt(15).GetSurfaceCell(
        constants.SurfaceColumn_Par1).DoubleValue = 2.0
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
    # print(spot_results.SpotData.GetReferenceCoordinate_X_For(1, 1))
    # print(spot_results.SpotData.GetReferenceCoordinate_Y_For(1, 1))

    # Zernike Standard Coefficients Analysis Results
    zernike = TheSystem.Analyses.New_Analysis(
        constants.AnalysisIDM_ZernikeStandardCoefficients)
    zernike_setting = zernike.GetSettings()
    baseSetting = CastTo(zernike_setting, 'IAS_ZernikeStandardCoefficients')
    baseSetting.Field.SetFieldNumber(0)
    baseSetting.Wavelength.UseAllWavelengths()
    baseSetting.Surface.UseImageSurface()

    # Extract Zernike Coefficients
    base = CastTo(zernike, 'IA_')
    base.ApplyAndWaitForCompletion()
    zernike_results = base.GetResults()
    res = CastTo(zernike_results, 'IAR_')
    print(res.GetTextFile(cwd + "\\hd.txt"))

    # analIDM = []
    # API_enum = list(constants.__dicts__[0].keys())
    # for i in API_enum:
    #     if i.find('AnalysisIDM_') != -1:
    #         analIDM.append(constants.__dicts__[0].get(i))
    #         print(constants.__dicts__[0].get(i), ':  ', i)
    # analIDM.sort()

    # # print('Name\tSetting\tDatGrid\tDatGridRgb\tDatSrs\tDatSrsRgb\t' +
    # #       'DatScat\tDatScatRgb\tRayData\tCriRayDat\tPathAnal\tSpotDat')

    # for k in analIDM:
    #     a = TheSystem.Analyses.New_Analysis(k)
    #     if a is None:
    #         print('This analysis cannot be opened in ',
    #               'Sequential Mode' if TheSystem.Mode == 0 else 'Non-Sequential Mode',
    #               ': enumID ', k)
    #         continue
    #     ar = a.GetResults()
    #     # print(a.GetAnalysisName, '\t',
    #     #     a.HasAnalysisSpecificSettings, '\t',
    #     #     ar.DataGrids is not None and ar.NumberOfDataGrids > 0, '\t',
    #     #     ar.DataGridsRgb is not None and ar.NumberOfDataGridsRgb > 0, '\t',
    #     #     ar.DataSeries is not None and ar.NumberOfDataSeries > 0, '\t',
    #     #     ar.DataSeriesRgb is not None and ar.NumberOfDataSeriesRgb > 0, '\t',
    #     #     ar.DataScatterPoints is not None and ar.NumberOfDataScatterPoints > 0, '\t',
    #     #     ar.DataScatterPointsRgb is not None and ar.NumberOfDataScatterPoints > 0, '\t',
    #     #     ar.RayData is not None, '\t',
    #     #     ar.CriticalRayData is not None, '\t',
    #     #     ar.PathAnalysisData is not None, '\t',
    #     #     ar.SpotData is not None)
    #     print(a.GetAnalysisName, '\t',
    #           a.HasAnalysisSpecificSettings, '\t',
    #           ar.HeaderData is not None, '\t',
    #           ar.MetaData is not None, '\t',
    #           ar.Messages is not None and ar.NumberOfMessages > 0)

    #     a.Close()

    # Set up Batch Ray Trace
    raytrace = TheSystem.Tools.OpenBatchRayTrace()
    nsur = TheSystem.LDE.NumberOfSurfaces
    max_rays = 30
    normUnPolData = raytrace.CreateNormUnpol(
        (max_rays + 1) * (max_rays + 1), constants.RaysType_Real, nsur)

    # Define batch ray trace constants
    hx = 0.0
    hy = 0.0
    max_wave = TheSystem.SystemData.Wavelengths.NumberOfWavelengths

    # Initialize x/y image plane arrays
    x_ary = np.empty((max_wave, ((max_rays + 1) * (max_rays + 1))))
    y_ary = np.empty((max_wave, ((max_rays + 1) * (max_rays + 1))))
    # print("x_ary shape", x_ary.shape)
    # print("x_ary shape", y_ary.shape)

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

    for wave in range(1, max_wave + 1):
        # Adding Rays to Batch
        normUnPolData.ClearData()
        for i in range(1, (max_rays + 1) * (max_rays + 1) + 1):

            px = np.random.random() * 2 - 1
            py = np.random.random() * 2 - 1

            while (px*px + py*py > 1):
                py = np.random.random() * 2 - 1
            normUnPolData.AddRay(
                wave, hx, hy, px, py, constants.OPDMode_None)

        baseTool = CastTo(raytrace, 'ISystemTool')
        baseTool.RunAndWaitForCompletion()

        # Read batch raytrace and display results
        normUnPolData.StartReadingResults()
        output = normUnPolData.ReadNextResult()

        while output[0]:    # success
            # ErrorCode & vignetteCode
            if ((output[2] == 0) and (output[3] == 0)):
                x_ary[wave - 1, output[1] - 1] = output[4]   # X
                y_ary[wave - 1, output[1] - 1] = output[5]   # Y
            output = normUnPolData.ReadNextResult()
        temp = plt.plot(np.squeeze(x_ary[wave - 1, :]), np.squeeze(
            y_ary[wave - 1, :]), '.', ms=3, c=colors[wave - 1],
            marker=markers[wave - 1])

    # np.savetxt("x.csv", x_ary, delimiter=",")
    # np.savetxt("y.csv", y_ary, delimiter=",")
    plt.title('Spot Diagram: %s' % (os.path.basename(testFile)))
    plt.draw()

    '''This will clean up the connection to OpticStudio.
    Note that it closes down the server instance of OpticStudio,
    so you for maximum performance do not do this until you need to.'''
    del zosapi
    zosapi = None

    # place plt.show() after clean up to release OpticStudio from memory
    # plt.show()
