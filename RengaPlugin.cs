using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;

namespace RengaPlugin
{
    public class RengaPlugin : Renga.IPlugin
    {
        private Renga.Application app;

        private string folderPath; 

        Dictionary<Guid, string> objectTypeNames;

        private Renga.ActionEventSource ElementsEventSource;
        private Renga.IAction ElementsAction;

        private Renga.ActionEventSource MaterialsEventSource;
        private Renga.IAction MaterialsAction;

        private Renga.ActionEventSource VolumesEventSource;
        private Renga.IAction VolumesAction;

        public bool Initialize(string pluginFolder)
        {
            app = new Renga.Application();
            folderPath = pluginFolder;
            Renga.IUI ui = app.UI;

            try
            {
                InitObjectTypeNames();
                InitButtons(ui);
            }
            catch (Exception ex)
            {
                ui.ShowMessageBox(Renga.MessageIcon.MessageIcon_Error, "Error", ex.Message);
                return false;
            }
            return true;
        }

        private void InitObjectTypeNames()
        {
            objectTypeNames = new Dictionary<Guid, string>
            {
                { Renga.ObjectTypes.Column,             "Колонна" },
                { Renga.ObjectTypes.Beam,               "Балка" },
                { Renga.ObjectTypes.WallFoundation,     "Ленточный фундамент" },
                { Renga.ObjectTypes.Plate,              "Плитный фундамент" },
                { Renga.ObjectTypes.IsolatedFoundation, "Столбчатый фундамент" },
                { Renga.ObjectTypes.Railing,            "Перила" },
                { Renga.ObjectTypes.Door,               "Дверь" },
                { Renga.ObjectTypes.Floor,              "Перекрытие" },
                { Renga.ObjectTypes.Ramp,               "Пандус" },
                { Renga.ObjectTypes.Opening,            "Проем" },
                { Renga.ObjectTypes.Roof,               "Крыша" },
                { Renga.ObjectTypes.Stair,              "Лестница" },
                { Renga.ObjectTypes.Wall,               "Стена" },
                { Renga.ObjectTypes.Window,             "Окно" }
            };
        }

        private void InitButtons(Renga.IUI ui)
        {
            Renga.IImage elIcon = ui.CreateImage();
            elIcon.LoadFromFile(Path.Combine(folderPath, @"Icons\elements.png"));

            Renga.IImage mtIcon = ui.CreateImage();
            mtIcon.LoadFromFile(Path.Combine(folderPath, @"Icons\materials.png"));

            Renga.IImage vlIcon = ui.CreateImage();
            vlIcon.LoadFromFile(Path.Combine(folderPath, @"Icons\volumes.png"));

            ElementsAction = ui.CreateAction();
            ElementsAction.DisplayName = "Перечень выбранных элементов";
            ElementsAction.Checkable = false;
            ElementsAction.Icon = elIcon;
            ElementsEventSource = new Renga.ActionEventSource(ElementsAction);
            ElementsEventSource.Triggered += ElementsAction_Triggered;

            MaterialsAction = ui.CreateAction();
            MaterialsAction.DisplayName = "Перечень материалов в проекте";
            MaterialsAction.Checkable = false;
            MaterialsAction.Icon = mtIcon;
            MaterialsEventSource = new Renga.ActionEventSource(MaterialsAction);
            MaterialsEventSource.Triggered += MaterialsAction_Triggered;

            VolumesAction = ui.CreateAction();
            VolumesAction.DisplayName = "Рассчитать объем в разрезе по материалам";
            VolumesAction.Checkable = false;
            VolumesAction.Icon = vlIcon;
            VolumesEventSource = new Renga.ActionEventSource(VolumesAction);
            VolumesEventSource.Triggered += VolumesAction_Triggered;

            Renga.IUIPanelExtension panel = ui.CreateUIPanelExtension();
            panel.AddToolButton(ElementsAction);
            panel.AddToolButton(MaterialsAction);
            panel.AddToolButton(VolumesAction);
            ui.AddExtensionToPrimaryPanel(panel);
        }

        private void ElementsAction_Triggered(object sender, EventArgs e)
        {
            Renga.ISelection selection = app.Selection;
            Array selectedObjects = selection.GetSelectedObjects();

            Dictionary<Guid, int> foundObjects = new Dictionary<Guid, int>();
            Renga.IModelObjectCollection objCollection = app.Project.Model.GetObjects();

            for (int i = 0; i < selectedObjects.GetLength(0); i++)
            {
                Renga.IModelObject obj = objCollection.GetById((int)selectedObjects.GetValue(i));

                if (foundObjects.ContainsKey(obj.ObjectType))
                    foundObjects[obj.ObjectType] += 1;
                else
                    foundObjects.Add(obj.ObjectType, 1);
            }

            StringBuilder builder = new StringBuilder();
            if (foundObjects.Count == 0)
            {
                app.UI.ShowMessageBox(Renga.MessageIcon.MessageIcon_Warning, 
                                      "Перечень выбранных элементов", 
                                      "Необходимо сначала выбрать элементы!");
                return;
            }
            foreach (KeyValuePair<Guid, int> item in foundObjects)
            {
                if (objectTypeNames.ContainsKey(item.Key))
                    builder.AppendLine($"{objectTypeNames[item.Key]}: {item.Value} ед.");
                else
                    continue;                    
            }
            app.UI.ShowMessageBox(Renga.MessageIcon.MessageIcon_Info, 
                                  "Перечень выбранных элементов", 
                                  builder.ToString());
        }

        private void MaterialsAction_Triggered(object sender, EventArgs e)
        {
            Renga.ILayeredMaterialManager materialManager = app.Project.LayeredMaterialManager;
            Renga.ILayeredMaterial material;

            StringBuilder sgLayerMessage = new StringBuilder();
            StringBuilder mlLayerMessage = new StringBuilder();
            
            for (int i = 0; i < 1000; i++)
            {
                if ((material = materialManager.GetLayeredMaterial(i)) == null) continue;
                if (material.Layers.Count == 1)
                    sgLayerMessage.AppendLine($"{material.Id}: {material.Name}");
                else
                    mlLayerMessage.AppendLine($"{material.Id}: {material.Name} (слоев: {material.Layers.Count} шт.)");
            }
            app.UI.ShowMessageBox(Renga.MessageIcon.MessageIcon_Info, "Однослойные материалы", sgLayerMessage.ToString());
            app.UI.ShowMessageBox(Renga.MessageIcon.MessageIcon_Info, "Многослойные материалы", mlLayerMessage.ToString());
        }

        private void VolumesAction_Triggered(object sender, EventArgs e)
        {
            try 
            {
                Renga.ISelection selection = app.Selection;
                List<Guid> allowedTypes = new List<Guid>()
                {
                    Renga.ObjectTypes.Floor,
                    Renga.ObjectTypes.Wall,
                    Renga.ObjectTypes.Roof,
                    Renga.ObjectTypes.Beam,
                    Renga.ObjectTypes.WallFoundation,
                    Renga.ObjectTypes.Plate,
                    Renga.ObjectTypes.IsolatedFoundation,
                    Renga.ObjectTypes.Ramp
                };
                Array selectedObjects = selection.GetSelectedObjects();
                Dictionary<int, double> materialVolumes = new Dictionary<int, double>();

                for (int i = 0; i < selectedObjects.GetLength(0); i++)
                {
                    Renga.IModelObjectCollection objCollection = app.Project.Model.GetObjects();
                    Renga.IModelObject obj = objCollection.GetById((int)selectedObjects.GetValue(i));
                    if (allowedTypes.Contains(obj.ObjectType))
                    {
                        Renga.IQuantity quantVolume = obj.GetQuantities().Get(Renga.QuantityIds.NetVolume);
                        double volumeMeters = quantVolume.AsVolume(Renga.VolumeUnit.VolumeUnit_Meters3);
                        Renga.IParameter param = obj.GetParameters().Get(Renga.ParameterIds.LayeredMaterialStyleId);
                        int materialID = param.GetIntValue();
                        if (materialVolumes.ContainsKey(materialID))
                            materialVolumes[materialID] += volumeMeters;
                        else
                            materialVolumes.Add(materialID, volumeMeters);
                    }
                    else continue;
                }

                Renga.ILayeredMaterialManager materialManager = app.Project.LayeredMaterialManager;
                if (materialVolumes.Count == 0)
                {
                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i < allowedTypes.Count; i++)
                        sb.AppendLine($"{i + 1}) {objectTypeNames[allowedTypes[i]]}");
                    app.UI.ShowMessageBox(Renga.MessageIcon.MessageIcon_Info,
                                          $"Объем элементов в разрезе по материалам",
                                          $"Среди выбранных объектов нет хотя бы одного из категории:" +
                                          Environment.NewLine + $"{sb}");
                    return;
                }
                
                StringBuilder builder = new StringBuilder();
                foreach (KeyValuePair<int, double> item in materialVolumes)
                {
                    Renga.ILayeredMaterial material = materialManager.GetLayeredMaterial(item.Key);
                    string materialName = (material == null) ? "не задан" : material.Name;
                    string materialVolume = $"{item.Value:#.000} куб. м";
                    builder.AppendLine($"Материал: {materialName}");
                    builder.AppendLine($"Объем: {materialVolume} {(materialVolumes.Last().Key == item.Key ? string.Empty : Environment.NewLine)}");
                }

                app.UI.ShowMessageBox(Renga.MessageIcon.MessageIcon_Info, "Объем объектов в разрезе по материалам", builder.ToString());
            }
            catch (Exception ex)
            {
                app.UI.ShowMessageBox(Renga.MessageIcon.MessageIcon_Error, "Renga Plugin", ex.Message);
            }
        }

        public void Stop()
        {
            ElementsEventSource.Dispose();
            MaterialsEventSource.Dispose();
            VolumesEventSource.Dispose();
        }
    }
}