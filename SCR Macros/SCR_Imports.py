import clr
#from collections import OrderedDict
import copy
import csv
from datetime import datetime, timedelta
from io import StringIO
import importlib
import json
import math
from multiprocessing import Pool, Process
import os
import random
import re
from re import search
import shutil
import subprocess
import sys
from threading import Lock, Thread
import time
from timeit import default_timer as timer
import uuid
import webbrowser
#import xml.etree.ElementTree as ET

clr.AddReference("IronPython.Wpf")
import wpf

clr.AddReference("MetadataExtractor")
from MetadataExtractor import Formats as MetadataFormats, GeoLocation as MetadataGeoLocation, ImageMetadataReader
from MetadataExtractor.Formats.Exif import ExifReader
from MetadataExtractor.Formats.Jpeg import JpegMetadataReader, JpegSegmentReader

import System
from System import Action, Array, Byte, Double, Environment, Func, Guid, Int32, IntPtr, Math, Reflection, String, Tuple, Type, UInt16, UInt32
from System.Collections.ObjectModel import ObservableCollection
from System.IO import File, MemoryStream, StreamReader
from System.Net import HttpStatusCode, HttpWebRequest, WebRequestMethods
from System.Reflection import BindingFlags
from System.Text import Encoding
from System.Windows import Application, FontWeights, GridLength, GridUnitType, MessageBox, MessageBoxButton, MessageBoxResult, PresentationSource, Visibility, Window
from System.Windows.Controls import ColumnDefinition, ComboBox, ComboBoxItem, DataGridTextColumn, Grid, GridViewColumn, ItemCollection, ListBox, ListBoxItem, ListView, \
                                    ListViewItem, RowDefinition, StackPanel, TextBlock, TextBox, TreeView, TreeViewItem
from System.Windows.Data import Binding
from System.Windows.Input import AccessKeyManager, Key, Keyboard, Mouse
from System.Windows.Media import VisualTreeHelper
from System.Windows.Media.Media3D import Point3D as SdePoint3D, Rect3D as SdeRect3D, Vector3D as SdeVector3D
from System.Windows.Shapes import Rectangle
from System.Windows.Threading import Dispatcher, DispatcherPriority

clr.AddReference("System.Core")
from System.Threading.Tasks import Task

clr.AddReference("System.Data")
from System.Data import DataColumn, DataRow, DataTable

clr.AddReference("System.Drawing")
from System.Drawing import Bitmap, Color, Image, Point as SystemDrawingPoint
from System.Drawing.Imaging import ImageFormat

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import BindingSource, Clipboard, ColumnHeader, Control, Cursor, DialogResult, FolderBrowserDialog, IMessageFilter, Keys, OpenFileDialog, SendKeys

clr.AddReference("System.Xml")
from System.Xml import XmlDocument

clr.AddReference("WindowsFormsIntegration")
from System.Windows.Forms.Integration import ElementHost

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
except:
    pass

from Trimble.Vce.Alignment import AlignmentLabel, DrapedLine, DynaView, Linestring, LocationComputer, ProfileView, Sideslope, SurfaceTie, VerticalAlignment

from Trimble.Vce.Alignment.Linestring import ElementFactory, IPointIdLocation, IXYZLocation, Linestring, VerticalElementFactory
try:
    from Trimble.Vce.Alignment.Linestring import LinestringHelper
except:
    LinestringHelper = None

from Trimble.Vce.Alignment.Linestring import ArcTangentType, IArcSegment, IBestFitArcSegment, IPIArcSegment, ISmoothCurveSegment, IStationVerticalLocation, IStraightSegment, \
                                             ITangentArcSegment, IVerticalCurve, IVerticalNoCurve

from Trimble.Vce.Collections import DynArray, Point3DArray

from Trimble.Vce.Coordinates import CoordComponentType, CoordinateSystemDefinition as CSD, CoordQuality, CoordSystem, CoordType, ICoordinate, KeyedIn, OfficeEnteredCoord, \
                                    Point as CoordPoint, PointCollection

from Trimble.Vce.Core import CalculateProjectEvents, DefaultAnalyticsReporter, EntityEventArgs, ModelEvents, TransactMethodCall

from Trimble.Vce.Core.Components import BlockView, EntityContainerBase, Layer, LayerCollection, LayerGroupCollection, LineStyleCollection, PointLabelStyle, Project, \
                                        SnapInAttributeExtension, TextStyle, TextStyleCollection, UserDefinedAttributes, ViewFilter, WorldView

from Trimble.Vce.Corridor import Template as CorridorTemplate

from Trimble.Vce.Data import FileProperties, FilePropertiesContainer, MediaFolder, MediaFolderContainer, PointCloudSelection, Shell3D, ShellMeshData, ShellMeshDataCollection, \
                             ShellMeshInstance
clr.AddReference ("Trimble.Vce.Data.GIS")  
from Trimble.Vce.Data.GIS import GisTransformer

from Trimble.Vce.Data.COGO import Calc

from Trimble.Vce.Data.Construction import TrimbleMap
from Trimble.Vce.Data.Construction.IFC import BIMEntity, BIMEntityCollection, IFCMesh # Shell3D

from Trimble.Vce.Data.Construction.Materials import MiscMaterial

from Trimble.Vce.Data.Construction.Settings import ConstructionCommandsSettings

from Trimble.Vce.Data.RawData import GeoReferencedImage, GeoReferencedImage3D, PointManager, RawDataBlock, RawDataContainer

clr.AddReference ("Trimble.Vce.Data.GeoRefImage")
from Trimble.Vce.Data.GeoRefImage import TiePoint, WorldFileFormat

clr.AddReference ("Trimble.Vce.Data.IFC")  
from Trimble.VCE.Data.IFC import IfcExporter

clr.AddReference ("Trimble.Vce.Data.RXL")
from Trimble.Vce.Data.RXL import FileWriter as RxlFileWriter, RXLAlignmentExporter, Versions as RxlVersions

clr.AddReference ("Trimble.Vce.Data.SDE")
from Trimble.Vce.Data.SDE import ColorizationState, GridScaledScan, Importer as SDEImporter
clr.AddReference ("SDE.NET")
from Trimble.Sde import Cloud as SdeCloud, CloudId as SdeCloudId, FilterByBox as SdeFilterByBox, FilterByRadius as SdeFilterByRadius, GutterPicker as SdeGutterPicker, \
                        List as SdeList, PointUV as SdePickerPointUV, ScanId as SdeScanId, SdeCloudOperator, SdeElevationType, SdePointInfo, SdePointSource, \
                        SpatialSamplingParameters as SdeSpatialSamplingParameters

from Trimble.Vce.Data.Scanning import DefaultExposedPointCloudRegion, ExposedPointCloudRegion, ExposedPointCloudRegionContainer, PointCloudDatabaseContainer, PointCloudRegion, \
                                      ScanningWorkflows

from Trimble.Vce.Features import Feature, FeatureManager, LineFeature, PointFeature
from Trimble.Vce.Features.FeatureCoding import FeatureCodeManager

from Trimble.Vce.ForeignCad import Arc as ArcObject, Circle as CadCircle, Face3D, Poly3D as CadPoly3D # we are also using Arc from geometry, PolyLine, PolyLineBase

from Trimble.Vce.ForeignCad import AttachmentPoint, BlockReference, DimArrowheadType, Hatch, Leader, LeaderType, MText, Point as CadPoint, PointLabelEntity as CadLabel, \
                                   Text as CadText, TextUtilities

from Trimble.Vce.Gem import CompositeSurface, ConstructionMaterialCollection, DisplayMode, DtmVolumes, ElevationSlopeTypes, Filer, Gem, GemMaterialMap, GemMaterials, GemVertexType, \
                            Model3D, Model3DCompSettings, Model3DContoursBuilder, Model3DQuickContours, ModelBoundaries, ProjectedSurface, SiteImprovementMaterialCollection, \
                            SlopingLevelSurface, SurfaceClassification, VextexAndTriangleList

clr.AddReference ("Trimble.Vce.Scanning")
from Trimble.Vce.Geometry import Point3DExtensions
from Trimble.Vce.Geometry import Arc, BestFitCircle, BiVector3D, Conversions, GeometryChainFitting, Intersections, Limits3D, Line3D, Matrix4D, Parabola, Plane3D, Point3D, PolySeg, \
                                 Primitive, PrimitiveLocation, RectangleD, SegmentFittingHelper, Side, Spinor3D, Triangle2D, Triangle3D, Vector2D, Vector3D

from Trimble.Vce.Geometry.PolySeg.Segment import AdjustmentType as SegmentAdjustmentType, Arc as ArcSegment, Line as SegmentLine, Parabola as ParabolaSegment, Segment, \
                                                 Type as SegmentType
try:
    #2023.10
    from Trimble.Vce.Geometry import FergusonSpline
except:
    #5.90
    from Trimble.Vce.Geometry.PolySeg.Segment import FergusonSpline


from Trimble.Vce.Geometry.Region import Region, RegionBuilder

from Trimble.Vce.Graphics import FillStyle, GraphicMarkerTypes, LeftMouseModeType, TrimbleColorCodeValues

clr.AddReference("Trimble.Vce.GraphicsEngine2D")
from Trimble.Vce.GraphicsEngine import OverlayBag
clr.AddReference("Trimble.Vce.Hoops")
from Trimble.Vce.GraphicsEngineHoops import PickAperature

from Trimble.Vce.Interfaces import Client, FeatureCoding, Point as PointHelper, ProgressBar
try: # new in 2024.00
    from Trimble.Vce.Interfaces import TransformationContext, TriangulationContext
except:
    pass

from Trimble.Vce.Interfaces.Client import CommandGranularity

from Trimble.Vce.Interfaces.Core import OverrideByLayer

try: #2024.00 and 2025.10?
    clr.AddReference ("Trimble.Sdk.CompositeEntities") 
    from Trimble.Sdk.CompositeEntities.Interfaces import ICompositeGeometry
except: #2025.21
    from Trimble.Vce.Interfaces.CompositeGeometry import ICompositeGeometry

try:
	#5.70
	from Trimble.Vce.Interfaces.Construction import UtilityNodeType
except:
    #5.60.2
    from Trimble.Vce.Utility import UtilityNodeType

from Trimble.Vce.Interfaces.SnapIn import DTMSharpness, ICollectionOfEntities, IHaveUcs, IMemberManagement, IName, IPointCloudReader, IPolyseg, ISnapIn, ISpacialContainer, SnapMode, \
                                          SurfaceRebuildMethod, TransformData

from Trimble.Vce.Interfaces.PointCloud import IExposedPointCloudRegion 
from Trimble.Vce.Interfaces.PointCloudIntegration import IPointCloudRegionIntegration

from Trimble.Vce.Interfaces.Units import AngularType, CoordinateType, ILinear, IParseProperties, LinearType

from Trimble.Vce.PlanSet import BasicSheet, PlanGridSheetSet, PlanSetSheetView, PlanSetSheetViews, ProfileSheetSet, SheetSet, SheetSetBase, XSSheetSet

from Trimble.Properties import PropertyKeys

clr.AddReference ("Trimble.Vce.Services.WeoGeo")
from Trimble.Vce.Services.TrimbleMaps import TileProperties, TileService, TrimbleMapsService, TrimbleMapsTileId

from Trimble.Vce.SurveyCAD import ProjectLineStyle, TextDisplayAttr

from Trimble.Vce.TTMetrics import StrokeFont, StrokeFontManager

from Trimble.Vce.Utility import UtilityNetwork, UtilityNode

clr.AddReference ("Trimble.Vce.UI.BaseCommands")
from Trimble.Vce.UI.BaseCommands import ExploreObjectControlHelper, SelectionContextMenuHandler, ViewHelper

clr.AddReference ("Trimble.Vce.UI.ConstructionCommands")
from Trimble.Vce.UI.ConstructionCommands import LimitSliceWindow
from Trimble.Vce.UI.ConstructionCommands.ProjectExplorer import BimEntityCollectionNode, BimEntityNode, BimGroupNode
from Trimble.Vce.UI.ConstructionCommands.ProjectExplorer2 import PlanSetCollection, PlanSetNode


clr.AddReference ("Trimble.Vce.UI.Controls")
from Trimble.Vce.UI.Controls import DisplayWindow, InputSettings, StationEdit, SurfaceTypeLists, TrimbleColor, VceMessageBox, Wpf as TBCWpf

clr.AddReference ("Trimble.Vce.UI.Client")
from Trimble.Vce.UI.ProjectExplorer2 import ExplorerData

# GlobalSelection moved to Trimble.Vce.Core in TBC 5.90
try:
    from Trimble.Vce.Core import GlobalSelection
except:
    from Trimble.Vce.UI.Controls import GlobalSelection
    
clr.AddReference("Trimble.Vce.UI.Hoops")
from Trimble.Vce.UI.Hoops import Hoops2dView, Hoops3dView, HoopsSheetView

clr.AddReference("Trimble.Vce.UI.ScanningCommands")
from Trimble.Vce.UI.ScanningCommands import ExtractionTools
try:
    from Trimble.Vce.UI.ScanningCommands import CreateBestFitLine
except:
    # moved to Trimble.Vce.Alignment.Linestring in TBC 5.90
    from Trimble.Vce.Alignment.Linestring import CreateBestFitLine

clr.AddReference("Trimble.Vce.UI.UIManager")
from Trimble.Vce.UI.UIManager import TrimbleOffice, UIEvents

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
    from Trimble.Sdk.UI import CursorStyle, DistanceType, I2DProjection, InputMethod, MousePosition, UIEventArgs, UndoWatcher
    from Trimble.Sdk.Interfaces import FeatureStatus
    from Trimble.Sdk.Interfaces.UI import ControlBoolean
except:
    try:
    	#5.50
        from Trimble.Sdk.Interfaces.UI import ControlBoolean, CursorStyle, DistanceType, I2DProjection, InputMethod, MousePosition, UIEventArgs
    except:
        try:
    		#5.40 or so
            from Trimble.Vce.UI.UIManager import UIEventArgs
            from Trimble.CustomControl.Interfaces import ControlBoolean, I2DProjection, MousePosition
            from Trimble.CustomControl.Interfaces.Enums import CursorStyle, InputMethod
        except:
    		# even older
            from Trimble.Vce.UI.Controls import ControlBoolean, CursorStyle, MousePosition, TrimbleColor

#sys.path.append("C:\Program Files\Sitech Construction Systems\ANZToolbox_24.0")
#clr.AddReference("Sitech.Data.D12D")
#from Sitech.Data.D12D.Exporter import D12daExporter, D12daExportBuilder
#from Sitech.Data.D12D import ProjectHeader as D12dProjectHeader

exec(open(r"C:\ProgramData\Trimble\MacroCommands3\SCR Macros\SCR_GlobalHelpers.py").read())





