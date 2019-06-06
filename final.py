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


class ZOSAPIAnalysis(object):

    def __init__(self, fileName):
        self.zosapi = PythonStandaloneApplication()
        value = self.zosapi.ExampleConstants()

        self.TheSystem = self.zosapi.TheSystem
        self.TheApplication = self.zosapi.TheApplication

        # Set up primary optical system
        import os
        cwd = os.getcwd()
        file = cwd + "\\" + fileName
        self.TheSystem.LoadFile(file, False)
        print("File Imported")

    def CreatePertubationFilter1(self, Decenter_X=0.0, Decenter_Y=0.0,
                                 Tilt_About_X=0.0, Tilt_About_Y=0.0,
                                 Tilt_About_Z=0.0):
        # Get Surfaces
        TheLDE = self.TheSystem.LDE
        SurfaceBefore = TheLDE.GetSurfaceAt(8)
        SurfaceAfter = TheLDE.GetSurfaceAt(11)

        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par1).DoubleValue = Decenter_X
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par2).DoubleValue = Decenter_Y
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par3).DoubleValue = Tilt_About_X
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par4).DoubleValue = Tilt_About_Y
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par5).DoubleValue = Tilt_About_Z

        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par1).DoubleValue = Decenter_X
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par2).DoubleValue = Decenter_Y
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par3).DoubleValue = Tilt_About_X
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par4).DoubleValue = Tilt_About_Y
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par5).DoubleValue = Tilt_About_Z

        print('Pertubation in Filter 1 Done!')

    def CreatePertubationFilter2(self, Decenter_X=0.0, Decenter_Y=0.0,
                                 Tilt_About_X=0.0, Tilt_About_Y=0.0,
                                 Tilt_About_Z=0.0):
        # Get Surfaces
        TheLDE = self.TheSystem.LDE
        SurfaceBefore = TheLDE.GetSurfaceAt(12)
        SurfaceAfter = TheLDE.GetSurfaceAt(15)

        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par1).DoubleValue = Decenter_X
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par2).DoubleValue = Decenter_Y
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par3).DoubleValue = Tilt_About_X
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par4).DoubleValue = Tilt_About_Y
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par5).DoubleValue = Tilt_About_Z

        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par1).DoubleValue = Decenter_X
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par2).DoubleValue = Decenter_Y
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par3).DoubleValue = Tilt_About_X
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par4).DoubleValue = Tilt_About_Y
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par5).DoubleValue = Tilt_About_Z

        print('Pertubation in Filter 2 Done!')

    def CreatePertubationLens(self, Decenter_X=0.0, Decenter_Y=0.0,
                              Tilt_About_X=0.0, Tilt_About_Y=0.0,
                              Tilt_About_Z=0.0):
        # Get Surfaces
        TheLDE = self.TheSystem.LDE
        SurfaceBefore = TheLDE.GetSurfaceAt(16)
        SurfaceAfter = TheLDE.GetSurfaceAt(19)

        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par1).DoubleValue = Decenter_X
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par2).DoubleValue = Decenter_Y
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par3).DoubleValue = Tilt_About_X
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par4).DoubleValue = Tilt_About_Y
        SurfaceBefore.GetSurfaceCell(
            constants.SurfaceColumn_Par5).DoubleValue = Tilt_About_Z

        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par1).DoubleValue = Decenter_X
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par2).DoubleValue = Decenter_Y
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par3).DoubleValue = Tilt_About_X
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par4).DoubleValue = Tilt_About_Y
        SurfaceAfter.GetSurfaceCell(
            constants.SurfaceColumn_Par5).DoubleValue = Tilt_About_Z

        print('Pertubation in Lens Done!')

    def ZernikeCoefficients(self, fileName=None):
        # Zernike Standard Coefficients Analysis Results
        zernike = self.TheSystem.Analyses.New_Analysis(
            constants.AnalysisIDM_ZernikeStandardCoefficients)
        zernike_setting = zernike.GetSettings()
        baseSetting = CastTo(
            zernike_setting, 'IAS_ZernikeStandardCoefficients')
        baseSetting.Field.SetFieldNumber(0)
        baseSetting.Wavelength.UseAllWavelengths()
        baseSetting.Surface.UseImageSurface()

        # Extract Zernike Coefficients
        base = CastTo(zernike, 'IA_')
        base.ApplyAndWaitForCompletion()
        zernike_results = base.GetResults()
        results = CastTo(zernike_results, 'IAR_')

        file = fileName if fileName is not None else "res.txt"
        import os
        cwd = os.getcwd()
        results.GetTextFile(cwd + "\\Zernike\\" + file)

    def CreateBatchRayTrace(self, max_rays=30):
        # Set up Batch Ray Trace
        raytrace = self.TheSystem.Tools.OpenBatchRayTrace()
        nsur = self.TheSystem.LDE.NumberOfSurfaces
        normUnPolData = raytrace.CreateNormUnpol(
            (max_rays + 1) * (max_rays + 1), constants.RaysType_Real, nsur)

        # Define batch ray trace constants
        hx = 0.0
        hy = 0.0
        max_wave = self.TheSystem.SystemData.Wavelengths.NumberOfWavelengths

        # Initialize x/y image plane arrays
        x_ary = np.empty((max_wave, ((max_rays + 1) * (max_rays + 1))))
        y_ary = np.empty((max_wave, ((max_rays + 1) * (max_rays + 1))))

        plt.rcParams["figure.figsize"] = (5, 5)
        colors = ('k', 'g', 'r')
        markers = ('+', 's', '^')

        if self.TheSystem.SystemData.Fields.GetFieldType() == constants.FieldType_Angle:
            field_type = 'Angle'
        elif self.TheSystem.SystemData.Fields.GetFieldType() == constants.FieldType_ObjectHeight:
            field_type = 'Height'
        elif self.TheSystem.SystemData.Fields.GetFieldType() == constants.FieldType_ParaxialImageHeight:
            field_type = 'Height'
        elif self.TheSystem.SystemData.Fields.GetFieldType() == constants.FieldType_RealImageHeight:
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
        plt.title('Spot Diagram')
        plt.draw()
        # plt.show()

    def SpotDiagramAnalysisResults(self):
        # Spot Diagram Analysis Results
        spot = self.TheSystem.Analyses.New_Analysis(
            constants.AnalysisIDM_StandardSpot)
        spot_setting = spot.GetSettings()
        baseSetting = CastTo(spot_setting, 'IAS_Spot')
        baseSetting.Field.SetFieldNumber(0)
        baseSetting.Wavelength.UseAllWavelengths()
        baseSetting.Surface.UseImageSurface()

        # extract RMS & Geo spot size for field points
        base = CastTo(spot, 'IA_')
        base.ApplyAndWaitForCompletion()
        spot_results = base.GetResults()
        print('RMS radius: %6.3f' %
              (spot_results.SpotData.GetRMSSpotSizeFor(1, 1)))
        print('GEO radius: %6.3f' %
              (spot_results.SpotData.GetGeoSpotSizeFor(1, 1)))


if __name__ == '__main__':
    file = "test2.zmx"
    ob = ZOSAPIAnalysis(file)

    ob.CreatePertubationFilter1()
    ob.SpotDiagramAnalysisResults()
    ob.CreateBatchRayTrace()
    ob.ZernikeCoefficients()

    '''This will clean up the connection to OpticStudio.
    Note that it closes down the server instance of OpticStudio,
    so you for maximum performance do not do this until you need to.'''
    del ob.zosapi
    ob.zosapi = None
