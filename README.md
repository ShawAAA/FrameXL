# FrameXL

Structural analysis in 2d for excel.

Created by Archie Shaw as a personal passion project. In the creation process of this tool I have used extremely minimal AI input (the exception being that AI was used to solve the beam actions/displacements equations for distributed moment loads and some general bugfixing advice). This tool has been developed exclusively by myself, and as such has not been subject to the level of review/testing a commerical-grade software would be subject to. Furthermore, the tool is in effective beta development currently, I intend to add workbooks that unit test the various functions of the tool highlighted below. While I have made reasonable (often painstaking) effort to eliminate bugs, I am certain that there are ways of using the tool or inherent issues within the tool itself that produce incorrect or misleading results. If in use of the tool you find any bugs (or if you have reasonable recommendations of future features) please document the issue and let me know on the Github or on Linkedin:

https://github.com/ShawAAA/FrameXL
www.linkedin.com/in/archie-shaw-0366441b9

Uses Euler-Bernoulli beams for structural analysis in 2d frames. Supports:
    
    Near exclusively cell based inputs, near exclusively cell based outputs. Allows direct connection of inputs/outputs to other calculations/sheets.
    Cell-note based input guidance.
    Designed to produce error values in cells if given misinputs rather than hanging your excel instance.
    One-button initialisation for new structures.
    Variable EA/EI properties.
    Beam end releases, spring end releases.
    Node springs (directionally translational and rotational).
    Nonlinear node springs, input via spring curve directionally.
    Point and distributed beam loads in x/y/zz directions in both local, global and projected coordinate systems.
    Assignment of loads to multiple nodes/elements by list, formula based load inputs (load calculated with respect to the position of the element/node in the coordinate system and member index with boolean logic allowed)
    Multiple non-sequential load case inputs, multiple loads per load case.
    Vehicle load creation wizard (converts chainage input and start node to a series of beam loads by sequential load cases).
    Additive and envelope load combination cases (where load combination cases can be added to and enveloped with other load combination cases).
    Load combination by list input (e.g. 1*LC1 to 1*LC10)
    Non-sequential load combinations (load combinations can reference combinations created in indexes after their own).
    Load combination envelopes preserving coexisting effects for peak cases.
    Graphical interface for viewing node and element loads/actions/reactions/displacements, by both load case and load case combination.
    Graphical interface is (approximately) to scale visually with automatic rescale and adjustable scaling of the viewed effect.
    Tabularised output of node and element loads/actions/reactions/displacements lists of load cases/combinations for lists of nodes/elements.
    Tabularised output contains envelope permutations by effect.
    Button to toggle recalculation off for performance when changing inputs for a complicated structure.
    Animation wizard to enable easy animation of graphs, generically written so it may be linked to excel cells in general (rather than just for specific applications in this tool).


To get started:

    Open the releases tab in Github and select the current release. 
    Download the applicable .XLL file and open it as an excel add-in. 
    Navigate to the "FrameXL" tab in the ribbon.
    Initialise a new structure (graphical representation is recommended for new users).
    Refer guidance on inputs in cell notes.


Planned features:

    Timoshenko-Ehrenfest beams, option to account for shear deformation of beams.
    Geometry creation wizard.
    Option to extract actions from an element at a given position rather than across the whole element.
    Nonlinear spring creation wizard.
    More descriptive in-cell error messages on calculation failure.
    Workbook unit tests of the features listed above.
    Buckling mode analysis.
    Improved guidance documentation.
    Nonlinear beam stiffnesses (currently not a priority).

Excluded features:

    3d analysis. While conceptually it is possible that 3d is not that much harder than 2d analysis on a software level, I dont believe I could ever do enough debugging/testing to have a sufficient degree of confidence that the results are correct. Furthermore, I would need to adopt an entirely different method of graphing the structure/results. Generally I see this tool as something to be used for fiddly small structures that are a bit too simple to be worth opening a more complicated software. Adding 3d functionality changes this usecase entirely.
    Capacity calculations/automatic component utilisation calculations. Ultimately I think this should be left to the user, the tool linking into user-created capacity sheets. I think that baking code clauses into the tool allows the user to design a structure while not understanding the underlying design assumptions. Furthermore it would increase the scope and complexity of the tool unacceptably.
