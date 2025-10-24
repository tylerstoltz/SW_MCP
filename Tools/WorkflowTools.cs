using System.ComponentModel;
using System.Runtime.InteropServices;
using ModelContextProtocol.Server;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksMCP.Services;

namespace SolidWorksMCP.Tools;

/// <summary>
/// Tier 1: High-level workflow tools for common SolidWorks operations
/// </summary>
[McpServerToolType]
public class WorkflowTools
{
    [McpServerTool]
    [Description("Creates a new part document in SolidWorks")]
    public static object CreatePart(
        ISolidWorksConnection connection,
        [Description("Template file to use (optional, uses default if not specified)")] string? templatePath = null)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            var app = connection.Application!;

            // Use default part template if none specified
            var template = templatePath ?? GetDefaultTemplate(app, (int)swDocumentTypes_e.swDocPART);

            // Create new part document
            var doc = (ModelDoc2?)app.NewDocument(template, 0, 0, 0);

            if (doc == null)
                return new { success = false, error = "Failed to create part document" };

            return new
            {
                success = true,
                documentType = "Part",
                documentName = doc.GetTitle(),
                path = doc.GetPathName()
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    [McpServerTool]
    [Description("Creates a new assembly document in SolidWorks")]
    public static object CreateAssembly(
        ISolidWorksConnection connection,
        [Description("Template file to use (optional, uses default if not specified)")] string? templatePath = null)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            var app = connection.Application!;
            var template = templatePath ?? GetDefaultTemplate(app, (int)swDocumentTypes_e.swDocASSEMBLY);

            var doc = (ModelDoc2?)app.NewDocument(template, 0, 0, 0);

            if (doc == null)
                return new { success = false, error = "Failed to create assembly document" };

            return new
            {
                success = true,
                documentType = "Assembly",
                documentName = doc.GetTitle(),
                path = doc.GetPathName()
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    [McpServerTool]
    [Description("Creates a sketch on a specified plane in the active document")]
    public static object CreateSketch(
        ISolidWorksConnection connection,
        [Description("Plane to create sketch on: Front, Top, Right, or a face/plane name")] string plane = "Front")
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            var doc = connection.ActiveDocument;
            if (doc == null)
                return new { success = false, error = "No active document" };

            var sketchMgr = doc.SketchManager;

            // Select the plane
            bool planeSelected = false;
            switch (plane.ToLowerInvariant())
            {
                case "front":
                    planeSelected = doc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;
                case "top":
                    planeSelected = doc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;
                case "right":
                    planeSelected = doc.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;
                default:
                    planeSelected = doc.Extension.SelectByID2(plane, "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;
            }

            if (!planeSelected)
                return new { success = false, error = $"Could not select plane: {plane}" };

            // Create the sketch
            sketchMgr.InsertSketch(true);

            return new
            {
                success = true,
                plane = plane,
                message = $"Sketch created on {plane} plane"
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    [McpServerTool]
    [Description("Creates an extruded boss/base or cut feature from the active sketch")]
    public static object CreateExtrude(
        ISolidWorksConnection connection,
        [Description("Extrusion depth in meters")] double depth,
        [Description("Reverse direction (default false)")] bool reverseDirection = false,
        [Description("Merge result with existing geometry (default true)")] bool mergeResult = true,
        [Description("Create a cut instead of boss/protrusion (default false)")] bool isCut = false,
        [Description("For sheet metal: cut normal to thickness (default false)")] bool normalCut = false)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            var doc = connection.ActiveDocument;
            if (doc == null)
                return new { success = false, error = "No active document" };

            var sketchMgr = doc.SketchManager;

            // Exit the sketch if we're in one
            if (sketchMgr.ActiveSketch != null)
            {
                doc.ClearSelection2(true);
                sketchMgr.InsertSketch(true);
            }

            // Select the sketch
            var featureMgr = doc.FeatureManager;

            Feature? feature;

            if (isCut)
            {
                // Create cut-extrude feature
                // FeatureCut4 signature: 27 parameters (SW 2017+)
                feature = (Feature?)featureMgr.FeatureCut4(
                    true,                                   // 1. Sd - Single-ended cut
                    false,                                  // 2. Flip - Don't flip side
                    reverseDirection,                       // 3. Dir - Reverse direction
                    (int)swEndConditions_e.swEndCondBlind,  // 4. T1 - End condition 1
                    0,                                      // 5. T2 - End condition 2 (not used for single-ended)
                    depth,                                  // 6. D1 - Depth for first end
                    0,                                      // 7. D2 - Depth for second end (not used)
                    false,                                  // 8. Dchk1 - No draft angle 1
                    false,                                  // 9. Dchk2 - No draft angle 2
                    false,                                  // 10. Ddir1 - Draft direction 1
                    false,                                  // 11. Ddir2 - Draft direction 2
                    0,                                      // 12. Dang1 - Draft angle 1
                    0,                                      // 13. Dang2 - Draft angle 2
                    false,                                  // 14. OffsetReverse1
                    false,                                  // 15. OffsetReverse2
                    false,                                  // 16. TranslateSurface1
                    false,                                  // 17. TranslateSurface2
                    normalCut,                              // 18. NormalCut - Cut normal to thickness (sheet metal)
                    false,                                  // 19. UseFeatScope
                    true,                                   // 20. UseAutoSelect
                    false,                                  // 21. AssemblyFeatureScope
                    false,                                  // 22. AutoSelectComponents
                    false,                                  // 23. PropagateFeatureToParts
                    (int)swStartConditions_e.swStartSketchPlane, // 24. T0 - Start condition
                    0,                                      // 25. StartOffset
                    false,                                  // 26. FlipStartOffset
                    false);                                 // 27. OptimizeGeometry
            }
            else
            {
                // Create boss/protrusion extrude feature
                // FeatureExtrusion2: 23 parameters (SW 2005+, superseded by FeatureExtrusion3 in SW 2014)
                feature = (Feature?)featureMgr.FeatureExtrusion2(
                    true,                                   // 1. Sd - Single-ended extrusion
                    false,                                  // 2. Flip - Don't flip side to cut
                    reverseDirection,                       // 3. Dir - Reverse direction flag
                    (int)swEndConditions_e.swEndCondBlind,  // 4. T1 - End condition for direction 1
                    0,                                      // 5. T2 - End condition for direction 2 (not used for single-ended)
                    depth,                                  // 6. D1 - Depth in meters for direction 1
                    0,                                      // 7. D2 - Depth for direction 2 (0 for single-ended)
                    false,                                  // 8. Dchk1 - No draft angle in direction 1
                    false,                                  // 9. Dchk2 - No draft angle in direction 2
                    false,                                  // 10. Ddir1 - Draft direction 1 (inward/outward)
                    false,                                  // 11. Ddir2 - Draft direction 2 (inward/outward)
                    0,                                      // 12. Dang1 - Draft angle for direction 1
                    0,                                      // 13. Dang2 - Draft angle for direction 2
                    false,                                  // 14. OffsetReverse1 - Offset direction for end 1
                    false,                                  // 15. OffsetReverse2 - Offset direction for end 2
                    false,                                  // 16. TranslateSurface1 - Translation vs offset for end 1
                    false,                                  // 17. TranslateSurface2 - Translation vs offset for end 2
                    mergeResult,                            // 18. Merge - Merge results in multibody part
                    false,                                  // 19. UseFeatScope - Feature affects selected bodies only
                    false,                                  // 20. UseAutoSelect - Auto-select all bodies
                    (int)swStartConditions_e.swStartSketchPlane, // 21. T0 - Start condition
                    0,                                      // 22. StartOffset - Offset from sketch plane
                    false);                                 // 23. FlipStartOffset - Flip start offset direction
            }

            if (feature == null)
                return new { success = false, error = $"Failed to create {(isCut ? "cut" : "boss")} extrude feature" };

            return new
            {
                success = true,
                featureName = feature.Name,
                featureType = isCut ? "cut" : "boss",
                depth = depth,
                message = $"{(isCut ? "Cut" : "Boss")} extrude feature created with depth {depth}m"
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    [McpServerTool]
    [Description("Adds a rectangular sketch to the active sketch")]
    public static object CreateRectangle(
        ISolidWorksConnection connection,
        [Description("X coordinate of first corner (meters)")] double x1,
        [Description("Y coordinate of first corner (meters)")] double y1,
        [Description("X coordinate of opposite corner (meters)")] double x2,
        [Description("Y coordinate of opposite corner (meters)")] double y2,
        [Description("Z coordinate (default 0)")] double z = 0)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            var doc = connection.ActiveDocument;
            if (doc == null)
                return new { success = false, error = "No active document" };

            var sketchMgr = doc.SketchManager;

            // Create corner rectangle
            sketchMgr.CreateCornerRectangle(x1, y1, z, x2, y2, z);

            return new
            {
                success = true,
                message = $"Rectangle created from ({x1}, {y1}) to ({x2}, {y2})"
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    [McpServerTool]
    [Description("Adds a circle to the active sketch")]
    public static object CreateCircle(
        ISolidWorksConnection connection,
        [Description("X coordinate of center (meters)")] double centerX,
        [Description("Y coordinate of center (meters)")] double centerY,
        [Description("Z coordinate of center (default 0)")] double centerZ,
        [Description("Radius of circle (meters)")] double radius)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            var doc = connection.ActiveDocument;
            if (doc == null)
                return new { success = false, error = "No active document" };

            var sketchMgr = doc.SketchManager;

            // Create circle by center and radius
            sketchMgr.CreateCircleByRadius(centerX, centerY, centerZ, radius);

            return new
            {
                success = true,
                message = $"Circle created at ({centerX}, {centerY}, {centerZ}) with radius {radius}m"
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    [McpServerTool]
    [Description("Adds a line to the active sketch")]
    public static object CreateLine(
        ISolidWorksConnection connection,
        [Description("X coordinate of start point (meters)")] double x1,
        [Description("Y coordinate of start point (meters)")] double y1,
        [Description("Z coordinate of start point (default 0)")] double z1,
        [Description("X coordinate of end point (meters)")] double x2,
        [Description("Y coordinate of end point (meters)")] double y2,
        [Description("Z coordinate of end point (default 0)")] double z2)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            var doc = connection.ActiveDocument;
            if (doc == null)
                return new { success = false, error = "No active document" };

            var sketchMgr = doc.SketchManager;

            // Create line segment
            var segment = sketchMgr.CreateLine(x1, y1, z1, x2, y2, z2);

            if (segment == null)
                return new { success = false, error = "Failed to create line" };

            return new
            {
                success = true,
                message = $"Line created from ({x1}, {y1}, {z1}) to ({x2}, {y2}, {z2})"
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    [McpServerTool]
    [Description("Gets information about the current model including features, sketches, and properties")]
    public static object GetModelInfo(ISolidWorksConnection connection)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            var doc = connection.ActiveDocument;
            if (doc == null)
                return new { success = false, error = "No active document" };

            var docType = (swDocumentTypes_e)doc.GetType();
            var features = new List<string>();
            var feature = (Feature?)doc.FirstFeature();

            while (feature != null)
            {
                features.Add($"{feature.Name} ({feature.GetTypeName()})");
                feature = (Feature?)feature.GetNextFeature();
            }

            return new
            {
                success = true,
                documentName = doc.GetTitle(),
                documentType = docType.ToString(),
                path = doc.GetPathName(),
                featureCount = features.Count,
                features = features.Take(20).ToArray(), // Limit to first 20
                isModified = doc.GetSaveFlag(),
                units = GetUnits(doc)
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    [McpServerTool]
    [Description("Saves the active document")]
    public static object SaveDocument(
        ISolidWorksConnection connection,
        [Description("File path to save to (optional, saves to current location if not specified)")] string? filePath = null)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            var doc = connection.ActiveDocument;
            if (doc == null)
                return new { success = false, error = "No active document" };

            int errors = 0, warnings = 0;
            bool result;

            if (!string.IsNullOrEmpty(filePath))
            {
                result = doc.Extension.SaveAs(filePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);
            }
            else
            {
                result = doc.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
            }

            if (!result)
                return new { success = false, error = $"Save failed. Errors: {errors}, Warnings: {warnings}" };

            return new
            {
                success = true,
                path = doc.GetPathName(),
                message = "Document saved successfully"
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    [McpServerTool]
    [Description("Creates a sheet metal base flange feature from the active sketch. IMPORTANT: Must have an ACTIVE sketch (not closed) - the method will automatically close it.")]
    public static object CreateBaseFlange(
        ISolidWorksConnection connection,
        [Description("Wall thickness in meters")] double thickness,
        [Description("Bend radius at corners in meters")] double bendRadius,
        [Description("Extrusion depth in meters")] double depth,
        [Description("Reverse extrusion direction (default false)")] bool reverseDirection = false,
        [Description("Direction to thicken: true for one side, false for midplane (default false)")] bool thickenOneDirection = false,
        [Description("Use default relief settings (default true)")] bool useDefaultRelief = true,
        [Description("Relief type: 0=Rectangular, 1=Tear, 2=Obround (default 0)")] int reliefType = 0,
        [Description("Relief width multiplier (default 1.0)")] double reliefRatio = 1.0,
        [Description("Merge with existing sheet metal bodies (default true)")] bool merge = true)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            var doc = connection.ActiveDocument;
            if (doc == null)
                return new { success = false, error = "No active document" };

            var sketchMgr = doc.SketchManager;

            // Check if there's an active sketch
            var activeSketch = sketchMgr.ActiveSketch;
            if (activeSketch == null)
                return new { success = false, error = "No active sketch found. Create a sketch first." };

            // Exit the sketch - use same pattern as CreateExtrude (which works)
            doc.ClearSelection2(true);
            sketchMgr.InsertSketch(true);

            var featureMgr = doc.FeatureManager;

            // Create base flange feature
            // InsertSheetMetalBaseFlange2 has 19 parameters (SW 2010+)
            // For sheet metal base flange as first feature, always use merge=false
            // (there are no existing sheet metal bodies to merge with)
            var shouldMerge = false;

            var feature = (Feature?)featureMgr.InsertSheetMetalBaseFlange2(
                thickness,                                  // 1. Thickness - Wall thickness
                thickenOneDirection,                        // 2. ThickenDir - Direction to thicken
                bendRadius,                                 // 3. Radius - Bend radius at corners
                depth,                                      // 4. ExtrudeDist1 - Extrusion distance direction 1
                0,                                          // 5. ExtrudeDist2 - Distance direction 2 (0 for single direction)
                reverseDirection,                           // 6. FlipExtruDir - Reverse extrusion direction
                (int)swEndConditions_e.swEndCondBlind,      // 7. EndCondition1 - Blind end condition
                (int)swEndConditions_e.swEndCondBlind,      // 8. EndCondition2 - Blind end condition
                0,                                          // 9. DirToUse - Direction to use (0=one direction)
                null,                                       // 10. PCBA - Custom bend allowance (null for default)
                useDefaultRelief,                           // 11. UseDefaultRelief - Use default relief settings
                reliefType,                                 // 12. ReliefType - Relief type (rectangular, tear, obround)
                0.001,                                      // 13. ReliefWidth - Relief width (1mm default)
                0.001,                                      // 14. ReliefDepth - Relief depth (1mm default)
                reliefRatio,                                // 15. ReliefRatio - Relief ratio
                false,                                      // 16. UseReliefRatio - Use width/depth, not ratio
                shouldMerge,                                // 17. Merge - Only merge if there are existing bodies
                false,                                      // 18. UseFeatScope - Use feature scope (false for all bodies)
                true);                                      // 19. UseAutoSelect - Auto select bodies

            if (feature == null)
                return new { success = false, error = "Failed to create base flange feature" };

            return new
            {
                success = true,
                featureName = feature.Name,
                thickness = thickness,
                bendRadius = bendRadius,
                depth = depth,
                message = $"Sheet metal base flange created with thickness {thickness}m and depth {depth}m"
            };
        }
        catch (Exception ex)
        {
            return new {
                success = false,
                error = ex.Message,
                exceptionType = ex.GetType().Name,
                stackTrace = ex.StackTrace
            };
        }
    }

    [McpServerTool]
    [Description("Creates an edge flange on selected edges of a sheet metal part")]
    public static object CreateEdgeFlange(
        ISolidWorksConnection connection,
        [Description("Method to select edge: 'current' (use currently selected), 'coordinates' (use x,y,z), or 'name' (try edge name)")] string selectionMethod = "current",
        [Description("Edge name to select (only used if selectionMethod='name')")] string? edgeName = null,
        [Description("X coordinate for edge selection in meters (only used if selectionMethod='coordinates')")] double x = 0,
        [Description("Y coordinate for edge selection in meters (only used if selectionMethod='coordinates')")] double y = 0,
        [Description("Z coordinate for edge selection in meters (only used if selectionMethod='coordinates')")] double z = 0,
        [Description("Flange length in meters")] double flangeLength = 0.025,
        [Description("Flange angle in degrees (default 90)")] double flangeAngle = 90,
        [Description("Bend radius in meters")] double bendRadius = 0.001,
        [Description("Bend position: 0=Material Inside, 1=Material Outside, 2=Bend Outside, 3=Bend From Virtual Sharp (default 0)")] int bendPosition = 0,
        [Description("Use default relief settings (default true)")] bool useDefaultRelief = true,
        [Description("Relief type: 0=Rectangular, 1=Tear, 2=Obround (default 0)")] int reliefType = 0)
    {
        try
        {
            if (!connection.IsConnected)
                return new { success = false, error = "Not connected to SolidWorks" };

            var doc = connection.ActiveDocument;
            if (doc == null)
                return new { success = false, error = "No active document" };

            // Handle edge selection based on method
            bool edgeSelected = false;
            var selMgr = (ISelectionMgr?)doc.SelectionManager;
            if (selMgr == null)
                return new { success = false, error = "Failed to get selection manager" };

            // Clear selections unless using current selection
            if (selectionMethod != "current")
            {
                doc.ClearSelection2(true);
            }

            switch (selectionMethod.ToLower())
            {
                case "current":
                    // Use the currently selected edge(s)
                    var count = selMgr.GetSelectedObjectCount2(-1);
                    if (count == 0)
                    {
                        return new {
                            success = false,
                            error = "No edges currently selected. Please select an edge first, or use 'coordinates' method with x,y,z values."
                        };
                    }
                    edgeSelected = true;
                    break;

                case "coordinates":
                    // Select edge by coordinates (most common in examples)
                    edgeSelected = doc.Extension.SelectByID2(
                        "",  // Empty string for name when using coordinates
                        "EDGE",
                        x, y, z,
                        false,
                        0,
                        null,
                        0);
                    if (!edgeSelected)
                    {
                        return new {
                            success = false,
                            error = $"Could not select edge at coordinates ({x}, {y}, {z}). Make sure coordinates point to an edge."
                        };
                    }
                    break;

                case "name":
                    if (string.IsNullOrEmpty(edgeName))
                    {
                        return new { success = false, error = "Edge name is required when using 'name' selection method" };
                    }
                    // Try different naming formats
                    string[] namesToTry = {
                        edgeName,                    // As provided
                        $"{edgeName}@Base-Flange1",  // With feature suffix
                        $"Edge<{edgeName}>",          // With angle brackets
                        $"Edge{edgeName}"             // Simple format
                    };

                    foreach (var name in namesToTry)
                    {
                        edgeSelected = doc.Extension.SelectByID2(
                            name,
                            "EDGE",
                            0, 0, 0,
                            false,
                            0,
                            null,
                            0);
                        if (edgeSelected) break;
                    }

                    if (!edgeSelected)
                    {
                        return new
                        {
                            success = false,
                            error = $"Could not select edge with name variations of '{edgeName}'. Try using 'coordinates' method or pre-select the edge and use 'current' method.",
                            triedNames = namesToTry
                        };
                    }
                    break;

                default:
                    return new { success = false, error = $"Invalid selection method '{selectionMethod}'. Use 'current', 'coordinates', or 'name'." };
            }

            // Get the selected edge
            // (selMgr already declared above)

            var selectedEdge = (Edge?)selMgr.GetSelectedObject6(1, -1);

            if (selectedEdge == null)
                return new { success = false, error = "Failed to get selected edge" };

            // IMPORTANT: The edge must remain selected for some versions of the API
            // Re-select the edge entity to ensure it stays selected
            var entity = (Entity?)selectedEdge;
            if (entity != null)
            {
                entity.Select4(false, null);
            }

            // Prepare edge array
            object[] edges = new object[] { selectedEdge };

            var featureMgr = doc.FeatureManager;

            // Convert angle from degrees to radians
            double angleRadians = flangeAngle * Math.PI / 180.0;

            // Calculate options based on whether custom values are provided
            // swInsertEdgeFlangeUseDefaultRadius = 1 (0x1)
            // swInsertEdgeFlangeUseDefaultRelief = 128 (0x80)
            int options = 0;
            double actualBendRadius = bendRadius;

            // Determine whether to use defaults based on parameters
            // If user provides specific values, don't use defaults even if useDefaultRelief is true
            bool shouldUseDefaults = useDefaultRelief && bendRadius == 0.001; // 0.001 is the default parameter value

            if (shouldUseDefaults)
            {
                // Use all defaults
                options = 1 + 128; // UseDefaultRadius + UseDefaultRelief
                actualBendRadius = 0;  // Set to 0 when using default radius
            }
            else
            {
                // Use custom values - user provided specific bend radius
                options = 0;
                actualBendRadius = bendRadius;
            }

            // Create edge flange feature
            // InsertSheetMetalEdgeFlange2 has 13 parameters (SW 2007+)
            // Based on Insert_Sheet_Metal_Edge_Flange_Example_VB.htm
            var feature = (Feature?)featureMgr.InsertSheetMetalEdgeFlange2(
                edges,                                      // 1. FlangeEdges - Array of edges
                null,                                       // 2. SketchFeats - Array of sketches (null for default profile)
                options,                                    // 3. BooleanOptions - Use defaults when requested
                angleRadians,                               // 4. FlangeAngle - Angle in radians
                actualBendRadius,                           // 5. FlangeRadius - Bend radius (0 when using default)
                bendPosition,                               // 6. BendPosition - Flange position type
                flangeLength,                               // 7. FlangeOffsetDist - Length of flange
                reliefType,                                 // 8. ReliefType - Relief type (ignored when using defaults)
                0,                                          // 9. FlangeReliefRatio - Set to 0 (not using ratio)
                0,                                          // 10. FlangeReliefWidth - Set to 0 when using defaults
                0,                                          // 11. FlangeReliefDepth - Set to 0 when using defaults
                0,                                          // 12. FlangeSharpType - Virtual sharp type (0 for default)
                null);                                      // 13. CustomBendAllowance - null for default

            if (feature == null)
            {
                // Add diagnostics to understand the failure
                var selCountAfter = selMgr.GetSelectedObjectCount2(-1);
                return new {
                    success = false,
                    error = "Failed to create edge flange feature - InsertSheetMetalEdgeFlange2 returned null",
                    diagnostics = new {
                        selectedCount = selCountAfter,
                        edgeWasFound = selectedEdge != null,
                        options = options,
                        angleRadians = angleRadians,
                        angleDegrees = flangeAngle,
                        bendRadius = actualBendRadius,
                        requestedBendRadius = bendRadius,
                        bendPosition = bendPosition,
                        flangeLength = flangeLength,
                        reliefType = reliefType,
                        useDefaultRelief = useDefaultRelief
                    }
                };
            }

            return new
            {
                success = true,
                featureName = feature.Name,
                edgeName = edgeName,
                flangeLength = flangeLength,
                flangeAngle = flangeAngle,
                bendRadius = bendRadius,
                message = $"Edge flange created on edge '{edgeName}' with length {flangeLength}m and angle {flangeAngle}°"
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    // Helper methods
    private static string GetDefaultTemplate(SldWorks app, int docType)
    {
        return app.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
    }

    private static string GetUnits(ModelDoc2 doc)
    {
        var userUnit = doc.Extension.GetUserPreferenceInteger(
            (int)swUserPreferenceIntegerValue_e.swUnitsLinear,
            (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);

        return userUnit switch
        {
            (int)swLengthUnit_e.swMM => "Millimeters",
            (int)swLengthUnit_e.swCM => "Centimeters",
            (int)swLengthUnit_e.swMETER => "Meters",
            (int)swLengthUnit_e.swINCHES => "Inches",
            (int)swLengthUnit_e.swFEET => "Feet",
            _ => "Unknown"
        };
    }
}
