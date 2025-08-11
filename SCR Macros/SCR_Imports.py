import clr
import copy
import csv
from datetime import timedelta, datetime
from io import StringIO
import json
import math
from multiprocessing import cpu_count, Process
import os
import random
import re
import shutil
import subprocess
import sys
import System
from threading import Thread, Lock
import time
from timeit import default_timer as timer
import webbrowser

#from collections import OrderedDict

clr.AddReference ('IronPython.Wpf')
import wpf

clr.AddReference ("MetadataExtractor")
from MetadataExtractor import ImageMetadataReader, GeoLocation as MetadataGeoLocation, Formats as MetadataFormats
from MetadataExtractor.Formats.Jpeg import JpegMetadataReader, JpegSegmentReader
from MetadataExtractor.Formats.Exif import ExifReader

from System import Math, Action, Func, String, Guid, Byte, UInt16, UInt32, Int32, Environment
from System import Type, Tuple, Array, Double
from System.Collections.ObjectModel import ObservableCollection
clr.AddReference ("System.Drawing")
from System.Drawing import Bitmap, Color
clr.AddReference ("System.Data")
from System.Data import DataTable, DataColumn, DataRow
from System.IO import StreamReader, MemoryStream, File
from System.Text import Encoding
from System.Windows import Window, Visibility, MessageBox, MessageBoxResult, MessageBoxButton, GridLength, GridUnitType
from System.Windows.Controls import StackPanel, ComboBox, ComboBoxItem, ListView, ListBox, ListBoxItem, ListViewItem, ItemCollection, TreeView, TreeViewItem, Grid, DataGridTextColumn, GridViewColumn, ColumnDefinition, RowDefinition, TextBlock, TextBox, ComboBox
from System.Windows.Data import Binding
clr.AddReference ("System.Windows.Forms")
from System.Windows.Forms import Clipboard, Control, Keys, SendKeys, OpenFileDialog, DialogResult, FolderBrowserDialog, ColumnHeader, BindingSource
clr.AddReference("WindowsFormsIntegration")
from System.Windows.Forms.Integration import ElementHost
from System.Windows.Input import Keyboard, AccessKeyManager
from System.Windows.Shapes import Rectangle
from System.Windows.Threading import Dispatcher
#import InfragisticsDlls
#from InfragisticsDlls import Workbook, Worksheet, WindowOptions, XamSpreadsheet

# we are referencing way too many Trimble functions
# but that way it is easier to fix in case they move stuff around again
try:
    clr.AddReference ("Trimble.Sdk") # In version 5.50, model assemblies are "pre-referenced"
    clr.AddReference ("Trimble.Sdk.Primitives") # contains Trimble.Vce.Collections
    clr.AddReference ("Trimble.Sdk.Features")
    clr.AddReference ("Trimble.Sdk.GraphicsModel") # Trimble.Vce.TTMetrics
    clr.AddReference ("Trimble.Sdk.Properties") # Trimble Properties
    clr.AddReference ("Trimble.Sdk.SurveyCad") # Trimble.Vce.SurveyCAD
    
except:
    clr.AddReference ("Trimble.DiskIO")
    clr.AddReference ("Trimble.Vce.Alignment")
    clr.AddReference ("Trimble.Vce.Collections")
    clr.AddReference ("Trimble.Vce.Coordinates")
    clr.AddReference ("Trimble.Vce.Core")
    clr.AddReference ("Trimble.Vce.Data")
    clr.AddReference ("Trimble.Vce.Data.COGO")
    clr.AddReference ("Trimble.Vce.Data.Construction")
    clr.AddReference ("Trimble.Vce.Data.Scanning")
    clr.AddReference ("Trimble.Vce.Data.RawData")
    clr.AddReference ("Trimble.Vce.Features")
    clr.AddReference ("Trimble.Vce.ForeignCad")
    clr.AddReference ("Trimble.Vce.Gem")
    clr.AddReference ("Trimble.Vce.Geometry")
    clr.AddReference ("Trimble.Vce.Graphics")
    clr.AddReference ("Trimble.Vce.Interfaces")
    clr.AddReference ("Trimble.Vce.TTMetrics")
    clr.AddReference ("Trimble.Vce.Units")

from Trimble.DiskIO import OptionsManager

try: # new in 2024.00
    clr.AddReference("Trimble.Rw.Core")
    from Trimble.Rw.Core.VectorMath.Geometry import Plane3D as RwPlane3D
    from Trimble.Rw.Core.VectorMath import Point3D as RwPoint3D
    clr.AddReference ("Trimble.Sdk.CompositeEntities") 
    from Trimble.Sdk.CompositeEntities.Interfaces import ICompositeGeometry
except:
    pass

from Trimble.Vce.Alignment import ProfileView, VerticalAlignment, Linestring, AlignmentLabel, DrapedLine, DynaView, SurfaceTie, Sideslope, LocationComputer

from Trimble.Vce.Alignment.Linestring import Linestring, ElementFactory, VerticalElementFactory, IXYZLocation, IPointIdLocation

from Trimble.Vce.Alignment.Linestring import IStraightSegment, ISmoothCurveSegment, IBestFitArcSegment, IArcSegment, IPIArcSegment, ITangentArcSegment, IStationVerticalLocation, IVerticalNoCurve, IVerticalCurve

from Trimble.Vce.Collections import Point3DArray, DynArray

from Trimble.Vce.Coordinates import Point as CoordPoint, ICoordinate, PointCollection, OfficeEnteredCoord, KeyedIn, CoordComponentType, CoordQuality, CoordSystem, CoordType, CoordinateSystemDefinition as CSD

from Trimble.Vce.Core import TransactMethodCall, ModelEvents, EntityEventArgs, CalculateProjectEvents, DefaultAnalyticsReporter

from Trimble.Vce.Core.Components import WorldView, Project, Layer, LayerGroupCollection, LayerCollection, TextStyle, TextStyleCollection, LineStyleCollection, BlockView, EntityContainerBase, PointLabelStyle, SnapInAttributeExtension, ViewFilter, UserDefinedAttributes

from Trimble.Vce.Corridor import Template as CorridorTemplate

from Trimble.Vce.Data import PointCloudSelection, FileProperties, FilePropertiesContainer, MediaFolder, MediaFolderContainer, Shell3D, ShellMeshDataCollection, ShellMeshInstance, ShellMeshData

from Trimble.Vce.Data.COGO import Calc

from Trimble.Vce.Data.Construction.IFC import BIMEntityCollection, BIMEntity, IFCMesh # Shell3D

from Trimble.Vce.Data.Construction.Materials import MiscMaterial

from Trimble.Vce.Data.Construction.Settings import ConstructionCommandsSettings

from Trimble.Vce.Data.RawData import PointManager, RawDataContainer, RawDataBlock, GeoReferencedImage, GeoReferencedImage3D

clr.AddReference ("Trimble.Vce.Data.GeoRefImage")
from Trimble.Vce.Data.GeoRefImage import WorldFileFormat, TiePoint

clr.AddReference ("Trimble.Vce.Data.IFC")  
from Trimble.VCE.Data.IFC import IfcExporter

clr.AddReference ("Trimble.Vce.Data.RXL")
from Trimble.Vce.Data.RXL import RXLAlignmentExporter, FileWriter as RxlFileWriter, Versions as RxlVersions

clr.AddReference ("Trimble.Vce.Data.SDE")
from Trimble.Vce.Data.SDE import ColorizationState, GridScaledScan, Importer as SDEImporter

from Trimble.Vce.Data.Scanning import PointCloudRegion, ExposedPointCloudRegion

from Trimble.Vce.Features import FeatureManager, LineFeature, PointFeature

from Trimble.Vce.ForeignCad import PolyLine, PolyLineBase, Poly3D, Arc as ArcObject, Circle as CadCircle, Face3D # we are also using Arc from geometry

from Trimble.Vce.ForeignCad import Point as CadPoint, MText, Text as CadText, BlockReference, Leader, LeaderType, DimArrowheadType, PointLabelEntity as CadLabel, AttachmentPoint, TextUtilities, Hatch

from Trimble.Vce.Gem import VextexAndTriangleList, Model3D, ProjectedSurface, SlopingLevelSurface, ElevationSlopeTypes, Gem, Filer, DtmVolumes, Model3DQuickContours, Model3DContoursBuilder, ModelBoundaries, GemVertexType, Model3DCompSettings, DisplayMode, SiteImprovementMaterialCollection, ConstructionMaterialCollection, GemMaterials, GemMaterialMap, SurfaceClassification

from Trimble.Vce.Geometry import Triangle2D, Triangle3D, Point3D, Plane3D, Arc, Matrix4D, Vector2D, Vector3D, BiVector3D, Spinor3D, PolySeg, Limits3D, Side, Intersections, RectangleD, Primitive, PrimitiveLocation, Conversions, Line3D, Parabola, BestFitCircle, GeometryChainFitting, SegmentFittingHelper

from Trimble.Vce.Geometry.PolySeg.Segment import Segment, Line as SegmentLine, Arc as ArcSegment, Parabola as ParabolaSegment, Type as SegmentType, AdjustmentType as SegmentAdjustmentType
try:
    #2023.10
    from Trimble.Vce.Geometry import FergusonSpline
except:
    #5.90
    from Trimble.Vce.Geometry.PolySeg.Segment import FergusonSpline


from Trimble.Vce.Geometry.Region import Region, RegionBuilder

from Trimble.Vce.Graphics import LeftMouseModeType, GraphicMarkerTypes, TrimbleColorCodeValues

clr.AddReference("Trimble.Vce.GraphicsEngine2D")
from Trimble.Vce.GraphicsEngine import OverlayBag

from Trimble.Vce.Interfaces import Client, ProgressBar, Point as PointHelper, FeatureCoding
try: # new in 2024.00
    from Trimble.Vce.Interfaces import TriangulationContext, TransformationContext
except:
    pass

from Trimble.Vce.Interfaces.Client import CommandGranularity

from Trimble.Vce.Interfaces.Core import OverrideByLayer

try:
	#5.70
	from Trimble.Vce.Interfaces.Construction import UtilityNodeType
except:
    #5.60.2
    from Trimble.Vce.Utility import UtilityNodeType

from Trimble.Vce.Interfaces.SnapIn import IPolyseg, IName, IHaveUcs, ISnapIn, IPointCloudReader, IMemberManagement, ICollectionOfEntities, TransformData, DTMSharpness, SurfaceRebuildMethod, SnapMode

from Trimble.Vce.Interfaces.PointCloud import IExposedPointCloudRegion
#from Trimble.Vce.Interfaces.PointCloudIntegration import LightWeightScanPointBatch

from Trimble.Vce.Interfaces.Units import ILinear, LinearType, CoordinateType

from Trimble.Vce.PlanSet import PlanSetSheetViews, PlanSetSheetView, BasicSheet, SheetSetBase, SheetSet, XSSheetSet, ProfileSheetSet, PlanGridSheetSet

from Trimble.Properties import PropertyKeys

from Trimble.Vce.SurveyCAD import ProjectLineStyle, TextDisplayAttr

from Trimble.Vce.TTMetrics import StrokeFont, StrokeFontManager

from Trimble.Vce.Utility import UtilityNode, UtilityNetwork

clr.AddReference ("Trimble.Vce.UI.BaseCommands")
from Trimble.Vce.UI.BaseCommands import ViewHelper, SelectionContextMenuHandler, ExploreObjectControlHelper

clr.AddReference ("Trimble.Vce.UI.ConstructionCommands")
from Trimble.Vce.UI.ConstructionCommands import LimitSliceWindow
from Trimble.Vce.UI.ConstructionCommands.ProjectExplorer import BimEntityCollectionNode, BimEntityNode, BimGroupNode 
from Trimble.Vce.UI.ConstructionCommands.ProjectExplorer2 import PlanSetCollection, PlanSetNode 


clr.AddReference ("Trimble.Vce.UI.Controls")
from Trimble.Vce.UI.Controls import SurfaceTypeLists, TrimbleColor, DisplayWindow, StationEdit, Wpf as TBCWpf, VceMessageBox, InputSettings

clr.AddReference ("Trimble.Vce.UI.Client")
from Trimble.Vce.UI.ProjectExplorer2 import ExplorerData

# GlobalSelection moved to Trimble.Vce.Core in TBC 5.90
try:
    from Trimble.Vce.Core import GlobalSelection
except:
    from Trimble.Vce.UI.Controls import GlobalSelection
    
clr.AddReference("Trimble.Vce.UI.Hoops")
from Trimble.Vce.UI.Hoops import Hoops2dView, Hoops3dView, HoopsSheetView

try:
    clr.AddReference("Trimble.Vce.UI.ScanningCommands")
    from Trimble.Vce.UI.ScanningCommands import CreateBestFitLine
except:
    # moved to Trimble.Vce.Alignment.Linestring in TBC 5.90
    from Trimble.Vce.Alignment.Linestring import CreateBestFitLine

clr.AddReference("Trimble.Vce.UI.UIManager")
from Trimble.Vce.UI.UIManager import UIEvents, TrimbleOffice

try: #5.90
    clr.AddReference("Trimble.Vce.ViewModel")
    from Trimble.Vce.ViewModel.CommandLine.CommandLineCommands import CommandHelper
except: #2023.10
    clr.AddReference("Trimble.Sdk.UI.Commands")
    from Trimble.Sdk.UI.Commands.CommandLine import CommandHelper

# later version of TBC moved these to a different assembly. If not found in new location, look at old
try:
    #2023.10
    #clr.AddReference("Trimble.Sdk.UI")
    from Trimble.Sdk.UI import MousePosition, InputMethod, CursorStyle, UIEventArgs, I2DProjection, DistanceType, UndoWatcher
    from Trimble.Sdk.Interfaces import FeatureStatus
    from Trimble.Sdk.Interfaces.UI import ControlBoolean
except:
    try:
    	#5.50
        from Trimble.Sdk.Interfaces.UI import InputMethod, MousePosition, CursorStyle, ControlBoolean, UIEventArgs, I2DProjection, DistanceType
    except:
        try:
    		#5.40 or so
            from Trimble.Vce.UI.UIManager import UIEventArgs
            from Trimble.CustomControl.Interfaces import MousePosition, ControlBoolean, I2DProjection
            from Trimble.CustomControl.Interfaces.Enums import CursorStyle, InputMethod
        except:
    		# even older
            from Trimble.Vce.UI.Controls import MousePosition, CursorStyle, ControlBoolean, TrimbleColor

#sys.path.append("C:\Program Files\Sitech Construction Systems\ANZToolbox_24.0")
#clr.AddReference("Sitech.Data.D12D")
#from Sitech.Data.D12D.Exporter import D12daExporter, D12daExportBuilder
#from Sitech.Data.D12D import ProjectHeader as D12dProjectHeader
