import clr
import sys
import csv
import copy
import math
import System
import json
import time
import shutil
import webbrowser
import re
from threading import Thread, Lock
from multiprocessing import cpu_count, Process
import random
import os
import subprocess

from timeit import default_timer as timer
from datetime import timedelta, datetime
from io import StringIO
#from collections import OrderedDict

clr.AddReference ('IronPython.Wpf')
import wpf

from System import Math, Action, Func, String, Guid, Byte, Int32, Environment
from System.Text import Encoding
from System.Windows import Window, Visibility, MessageBox, MessageBoxResult, MessageBoxButton
from System.Windows.Controls import StackPanel, ComboBox, ComboBoxItem, ListBox, ListBoxItem, ItemCollection, TreeView, TreeViewItem, Grid
from System.Windows.Input import Keyboard, AccessKeyManager
from System import Type, Tuple, Array, Double
from System.IO import StreamReader, MemoryStream, File
from System.Windows.Threading import Dispatcher
clr.AddReference ("System.Drawing")
from System.Drawing import Bitmap, Color
clr.AddReference ("System.Windows.Forms")
from System.Windows.Forms import Control, Keys, OpenFileDialog, DialogResult, FolderBrowserDialog, ColumnHeader
#import InfragisticsDlls
#from InfragisticsDlls import Workbook, Worksheet, WindowOptions, XamSpreadsheet

# we are referencing way too many Trimble functions
# but that way it is easier to fix in case they move stuff around again
try:
    clr.AddReference ("Trimble.Sdk") # In version 5.50, model assemblies are "pre-referenced"
    clr.AddReference ("Trimble.Sdk.Primitives") # contains Trimble.Vce.Collections
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
    clr.AddReference ("Trimble.Vce.ForeignCad")
    clr.AddReference ("Trimble.Vce.Gem")
    clr.AddReference ("Trimble.Vce.Geometry")
    clr.AddReference ("Trimble.Vce.Graphics")
    clr.AddReference ("Trimble.Vce.Interfaces")
    clr.AddReference ("Trimble.Vce.TTMetrics")
    clr.AddReference ("Trimble.Vce.Units")

from Trimble.DiskIO import OptionsManager

from Trimble.Vce.Alignment import ProfileView, VerticalAlignment, Linestring, AlignmentLabel, DrapedLine, DynaView, SurfaceTie

from Trimble.Vce.Alignment.Linestring import Linestring, ElementFactory, IXYZLocation, IPointIdLocation

from Trimble.Vce.Alignment.Linestring import IStraightSegment, ISmoothCurveSegment, IBestFitArcSegment, IArcSegment, IPIArcSegment, ITangentArcSegment

from Trimble.Vce.Collections import Point3DArray, DynArray

from Trimble.Vce.Coordinates import Point as CoordPoint, ICoordinate, PointCollection, OfficeEnteredCoord, KeyedIn, CoordComponentType, CoordQuality, CoordSystem, CoordType

from Trimble.Vce.Core import TransactMethodCall, ModelEvents

from Trimble.Vce.Core.Components import WorldView, Project, Layer, LayerGroupCollection, LayerCollection, TextStyle, TextStyleCollection, LineStyleCollection, BlockView, EntityContainerBase, PointLabelStyle, SnapInAttributeExtension, ViewFilter

from Trimble.Vce.Data import PointCloudSelection, FileProperties, FilePropertiesContainer, MediaFolder, MediaFolderContainer

from Trimble.Vce.Data.COGO import Calc

from Trimble.Vce.Data.Construction.IFC import BIMEntity, IFCMesh # Shell3D

from Trimble.Vce.Data.Construction.Materials import MiscMaterial

from Trimble.Vce.Data.Construction.Settings import ConstructionCommandsSettings

from Trimble.Vce.Data.RawData import PointManager, RawDataContainer, RawDataBlock, GeoReferencedImage

clr.AddReference ("Trimble.Vce.Data.RXL")
from Trimble.Vce.Data.RXL import RXLAlignmentExporter, FileWriter as RxlFileWriter, Versions as RxlVersions

clr.AddReference ("Trimble.Vce.Data.GeoRefImage")
from Trimble.Vce.Data.GeoRefImage import WorldFileFormat, TiePoint

clr.AddReference ("Trimble.Vce.Data.IFC")  
from Trimble.VCE.Data.IFC import IfcExporter

clr.AddReference ("Trimble.Vce.Data.RXL")
from Trimble.Vce.Data.RXL import RXLAlignmentExporter, FileWriter as RxlFileWriter, Versions as RxlVersions

from Trimble.Vce.Data.Scanning import PointCloudRegion, ExposedPointCloudRegion

from Trimble.Vce.ForeignCad import PolyLine, Poly3D, Arc as ArcObject, Circle as CadCircle, Face3D # we are also using Arc from geometry

from Trimble.Vce.ForeignCad import Point as CadPoint, MText, Text as CadText, BlockReference, Leader, LeaderType, DimArrowheadType, PointLabelEntity as CadLabel, AttachmentPoint

from Trimble.Vce.Gem import VextexAndTriangleList, Model3D, ProjectedSurface, Gem, Filer, DtmVolumes, Model3DQuickContours, Model3DContoursBuilder, ModelBoundaries, GemVertexType, Model3DCompSettings, DisplayMode, SiteImprovementMaterialCollection, ConstructionMaterialCollection, GemMaterials, GemMaterialMap, SurfaceClassification

from Trimble.Vce.Geometry import Triangle2D, Triangle3D, Point3D, Plane3D, Arc, Matrix4D, Vector2D, Vector3D, BiVector3D, Spinor3D, PolySeg, Limits3D, Side, Intersections, RectangleD, Primitive, PrimitiveLocation, Conversions, Line3D, Parabola

from Trimble.Vce.Geometry.PolySeg.Segment import Segment, Line as SegmentLine, Arc as ArcSegment, Parabola as ParabolaSegment
try:
    #2023.10
    from Trimble.Vce.Geometry import FergusonSpline
except:
    #5.90
    from Trimble.Vce.Geometry.PolySeg.Segment import FergusonSpline


from Trimble.Vce.Geometry.Region import Region, RegionBuilder

from Trimble.Vce.Graphics import LeftMouseModeType, GraphicMarkerTypes

clr.AddReference("Trimble.Vce.GraphicsEngine2D")
from Trimble.Vce.GraphicsEngine import OverlayBag

from Trimble.Vce.Interfaces import Client, ProgressBar, Point as PointHelper

from Trimble.Vce.Interfaces.Client import CommandGranularity

from Trimble.Vce.Interfaces.Core import OverrideByLayer

try:
	#5.70
	from Trimble.Vce.Interfaces.Construction import UtilityNodeType
except:
    #5.60.2
    from Trimble.Vce.Utility import UtilityNodeType

from Trimble.Vce.Interfaces.SnapIn import IPolyseg, IName, IHaveUcs, ISnapIn, IPointCloudReader, IMemberManagement, ICollectionOfEntities, TransformData, DTMSharpness

from Trimble.Vce.Interfaces.PointCloud import IExposedPointCloudRegion

from Trimble.Vce.Interfaces.Units import ILinear, LinearType, CoordinateType

from Trimble.Vce.PlanSet import PlanSetSheetViews, PlanSetSheetView, SheetSet, BasicSheet

from Trimble.Properties import PropertyKeys

from Trimble.Vce.SurveyCAD import ProjectLineStyle, TextDisplayAttr

from Trimble.Vce.TTMetrics import StrokeFont, StrokeFontManager

from Trimble.Vce.Utility import UtilityNode, UtilityNetwork

clr.AddReference ("Trimble.Vce.UI.BaseCommands")
from Trimble.Vce.UI.BaseCommands import ViewHelper, SelectionContextMenuHandler, ExploreObjectControlHelper

clr.AddReference ("Trimble.Vce.UI.ConstructionCommands")
from Trimble.Vce.UI.ConstructionCommands import LimitSliceWindow

clr.AddReference ("Trimble.Vce.UI.Controls")
from Trimble.Vce.UI.Controls import SurfaceTypeLists, TrimbleColor, ExplorerUI, ExplorerItemCollection, DisplayWindow, StationEdit, Wpf as TBCWpf, VceMessageBox

# GlobalSelection moved to Trimble.Vce.Core in TBC 5.90
try:
    from Trimble.Vce.Core import GlobalSelection
except:
    from Trimble.Vce.UI.Controls import GlobalSelection
    
clr.AddReference("Trimble.Vce.UI.Hoops")
from Trimble.Vce.UI.Hoops import Hoops2dView, Hoops3dView

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
    from Trimble.Sdk.UI import MousePosition, InputMethod, CursorStyle, UIEventArgs, I2DProjection, DistanceType
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
