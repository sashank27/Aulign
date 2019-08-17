from win32com.client.gencache import EnsureDispatch, EnsureModule
from win32com.client import CastTo, constants
import os
import matplotlib.pyplot as plt
import numpy as np
import inspect

# Function to extract Zernike coefficients from the results file obtained from the OpticStudio
def extractZernikeCoefficents(file):
        import codecs
        import re

        pattern = "^Z\s*(\d*)\s*([-]?[0-9]+[,.]?[0-9]*).*$"

        x = []
        for line in codecs.open(file, 'r', encoding='utf16'):
            match = re.search(pattern, line)
            if match is not None:
                x.append(float(match.group(2)))

        coefficients = np.array(x)
        return coefficients


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

class Parameter(object):

    def __init__(self, name, optimal, minval, maxval, step, optimal1=None):
        self.name = name
        self.optimal = optimal
        self.minval = minval
        self.maxval = maxval
        self.step = step
        
        if optimal1 is not None:
            self.optimal1 = optimal1

class Component(object):
    
    def __init__(self, surface, revSurface, compName):
        self.surface = surface
        self.compName = compName
        if revSurface is not None:
            self.revSurface = revSurface
    
    def createParameter(self, name, optimal, minval, maxval, step, optimal1=None):
        
        param = Parameter(name, optimal, minval, maxval, step, optimal1)
        setattr(self, name, param)
        print(name, param)

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

    def definePertubationAndGetResults(self, component, params=["dx"]):
        import os
        cwd = os.getcwd()

        param1 = component.params[0]
        vals1 = np.linspace(param1.minval, param1.maxval, num=((param1.maxval-param1.minval)/param1.step)+1)
        param2 = component.params[1]
        vals2 = np.linspace(param2.minval, param2.maxval, num=((param2.maxval-param2.minval)/param2.step)+1)
        
        for i in vals1:
            if i == 0:
                continue
            
            for j in vals2:
                if j == 0:
                    continue
                print(param1.name + " = " + "{0:.3f}".format(i) + " " + param2.name + " = " + "{0:.3f}".format(j))
                self.datadir = cwd + "\\data\\" + component.compName + "\\" + param1.name + "{0:.3f}".format(i) + param2.name + "{0:.3f}".format(j) + "\\"
                if not os.path.exists(self.datadir):
                    os.makedirs(self.datadir)

                Decenter_X = Decenter_Y = Tilt_About_X = Tilt_About_Y = Tilt_About_Z = 0

                if param == "dx":
                    Decenter_X = optimum1 + i
                if param == "dy":
                    Decenter_Y = optimum1 + i
                if param == "tx":
                    Tilt_About_X = optimum1 + i
                if param == "ty":
                    Tilt_About_Y = optimum1 + i
                if param == "tz":
                    Tilt_About_Z = optimum1 + i

                self.CreatePertubationInSurface(surfaceNumber, Decenter_X=Decenter_X, Decenter_Y=Decenter_Y,
                                                Tilt_About_X=Tilt_About_X, Tilt_About_Y=Tilt_About_Y, Tilt_About_Z=Tilt_About_Z)

                if revSurfaceNumber is not None:
                    Decenter_X = Decenter_Y = Tilt_About_X = Tilt_About_Y = Tilt_About_Z = 0

                    if param == "dx":
                        Decenter_X = optimum2 - i
                    if param == "dy":
                        Decenter_Y = optimum2 - i
                    if param == "tx":
                        Tilt_About_X = optimum2 - i
                    if param == "ty":
                        Tilt_About_Y = optimum2 - i
                    if param == "tz":
                        Tilt_About_Z = optimum2 - i

                    self.CreatePertubationInSurface(revSurfaceNumber, Decenter_X=Decenter_X, Decenter_Y=Decenter_Y,
                                                    Tilt_About_X=Tilt_About_X, Tilt_About_Y=Tilt_About_Y, Tilt_About_Z=Tilt_About_Z)

                self.SpotDiagramAnalysisResults()
                self.ZernikeCoefficients()
                # self.CreateBatchRayTrace()
                # print("aaa")

    def createComponents(self):
        TheLDE = self.TheSystem.LDE
        
        # Primary mirror
        Surface1 = TheLDE.GetSurfaceAt(5)
        Surface2 = TheLDE.GetSurfaceAt(7)

        self.pm = Component(Surface1, Surface2)
        self.pm.createParameter("dx", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par1).DoubleValue, -0.05, 0.05, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par1).DoubleValue)
        self.pm.createParameter("dy", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par2).DoubleValue, -0.05, 0.05, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par2).DoubleValue)
        self.pm.createParameter("tx", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par3).DoubleValue, -0.017, 0.017, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par3).DoubleValue)
        self.pm.createParameter("ty", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par4).DoubleValue, -0.017, 0.017, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par4).DoubleValue)

        # Secondary mirror
        Surface1 = TheLDE.GetSurfaceAt(8)
        Surface2 = TheLDE.GetSurfaceAt(10)

        self.sm = Component(Surface1, Surface2)
        self.sm.createParameter("dx", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par1).DoubleValue, -0.05, 0.05, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par1).DoubleValue)
        self.sm.createParameter("dy", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par2).DoubleValue, -0.05, 0.05, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par2).DoubleValue)
        self.sm.createParameter("tx", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par3).DoubleValue, -0.017, 0.017, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par3).DoubleValue)
        self.sm.createParameter("ty", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par4).DoubleValue, -0.017, 0.017, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par4).DoubleValue)

        # Lens
        Surface1 = TheLDE.GetSurfaceAt(20)
        Surface2 = TheLDE.GetSurfaceAt(23)

        self.lens = Component(Surface1, Surface2)
        self.lens.createParameter("dx", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par1).DoubleValue, -0.02, 0.02, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par1).DoubleValue)
        self.lens.createParameter("dy", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par2).DoubleValue, -0.02, 0.02, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par2).DoubleValue)
        self.lens.createParameter("tx", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par3).DoubleValue, -0.017, 0.017, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par3).DoubleValue)
        self.lens.createParameter("ty", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par4).DoubleValue, -0.017, 0.017, 0.001, Surface2.GetSurfaceCell(constants.SurfaceColumn_Par4).DoubleValue)

        # CCD
        Surface1 = TheLDE.GetSurfaceAt(23)
        
        self.ccd = Component(Surface1)
        self.ccd.createParameter("dx", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par1).DoubleValue, -2, 2, 0.001)
        self.ccd.createParameter("dy", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par2).DoubleValue, -0.05, 0.05, 0.001)
        self.ccd.createParameter("ty", Surface1.GetSurfaceCell(constants.SurfaceColumn_Par3).DoubleValue, -0.05, 0.05, 0.001)

    def CreatePertubationInSurface(self, surfaceNumber, Decenter_X=0.0, Decenter_Y=0.0,
                                   Tilt_About_X=0.0, Tilt_About_Y=0.0,
                                   Tilt_About_Z=0.0):
        # Get Surfaces
        TheLDE = self.TheSystem.LDE
        Surface = TheLDE.GetSurfaceAt(surfaceNumber)

        Surface.GetSurfaceCell(
            constants.SurfaceColumn_Par1).DoubleValue = float(Decenter_X)
        Surface.GetSurfaceCell(
            constants.SurfaceColumn_Par2).DoubleValue = float(Decenter_Y)
        Surface.GetSurfaceCell(
            constants.SurfaceColumn_Par3).DoubleValue = float(Tilt_About_X)
        Surface.GetSurfaceCell(
            constants.SurfaceColumn_Par4).DoubleValue = float(Tilt_About_Y)
        Surface.GetSurfaceCell(
            constants.SurfaceColumn_Par5).DoubleValue = float(Tilt_About_Z)

        # print('Pertubation in Surface {} Done!'.format(surfaceNumber))

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
        path = self.datadir + file
        results.GetTextFile(path)
        coefficients = extractZernikeCoefficents(path)
        np.savetxt(self.datadir + "zernike.csv", coefficients, delimiter=",")
        zernike.Close()

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

        np.savetxt(self.datadir + "x.csv", x_ary, delimiter=",")
        np.savetxt(self.datadir + "y.csv", y_ary, delimiter=",")
        plt.title('Spot Diagram')
        plt.savefig(self.datadir + "SpotDiagram.png")
        raytrace.Close()

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
        f = open(self.datadir + "rmsgeo.txt", "w+")
        f.write(str(spot_results.SpotData.GetRMSSpotSizeFor(1, 1)) + "\n" +
                str(spot_results.SpotData.GetGeoSpotSizeFor(1, 1)))
        # print('RMS radius: %6.3f' %
        #       (spot_results.SpotData.GetRMSSpotSizeFor(1, 1)))
        # print('GEO radius: %6.3f' %
        #       (spot_results.SpotData.GetGeoSpotSizeFor(1, 1)))
        spot.Close()


if __name__ == '__main__':
    file = "test2.zmx"
    ob = ZOSAPIAnalysis(file)

    ob.createComponents()

    # # Perturbation in Primary Mirror (Surface-5,7)
    # print('Perturbation in Primary Mirror')
    # ob.definePertubationAndGetResults(
    #     5, 7, element="pm", param="dx", minval=-0.05, maxval=0.05, step=0.001)
    # ob.definePertubationAndGetResults(
    #     5, 7, element="pm", param="dy", minval=-0.05, maxval=0.05, step=0.001)
    # ob.definePertubationAndGetResults(
    #     5, 7, element="pm", param="tx", minval=-0.017, maxval=0.017, step=0.001)
    # ob.definePertubationAndGetResults(
    #     5, 7, element="pm", param="ty", minval=-0.017, maxval=0.017, step=0.001)

    # # Perturbation in Secondary Mirror (Surface-8,10)
    # print('Perturbation in Secondary Mirror')
    # ob.definePertubationAndGetResults(
    #     8, 10, element="sm", param="dx", minval=-0.05, maxval=0.05, step=0.001)
    # ob.definePertubationAndGetResults(
    #     8, 10, element="sm", param="dy", minval=-0.05, maxval=0.05, step=0.001)
    # ob.definePertubationAndGetResults(
    #     8, 10, element="sm", param="tx", minval=-0.017, maxval=0.017, step=0.001)
    # ob.definePertubationAndGetResults(
    #     8, 10, element="sm", param="ty", minval=-0.017, maxval=0.017, step=0.001)

    # # Perturbation in Lens (Surface-20,23)
    # print('Perturbation in Lens')
    # ob.definePertubationAndGetResults(
    #     20, 23, element="lens", param="dx", minval=-0.02, maxval=0.02, step=0.001)
    # ob.definePertubationAndGetResults(
    #     20, 23, element="lens", param="dy", minval=-0.02, maxval=0.02, step=0.001)
    # ob.definePertubationAndGetResults(
    #     20, 23, element="lens", param="tx", minval=-0.017, maxval=0.017, step=0.001)
    # ob.definePertubationAndGetResults(
    #     20, 23, element="lens", param="ty", minval=-0.017, maxval=0.017, step=0.001)

    # # Perturbation in CCD (Surface-23)
    # print('Perturbation in CCD')
    # ob.definePertubationAndGetResults(
    #     23, element="ccd", param="dx", minval=-2, maxval=2, step=0.01)
    # ob.definePertubationAndGetResults(
    #     23, element="ccd", param="dy", minval=-2, maxval=2, step=0.01)
    # ob.definePertubationAndGetResults(
    #     23, element="ccd", param="ty", minval=-0.05, maxval=0.05, step=0.001)

    '''This will clean up the connection to OpticStudio.
    Note that it closes down the server instance of OpticStudio,
    so you for maximum performance do not do this until you need to.'''
    del ob.zosapi
    ob.zosapi = None
