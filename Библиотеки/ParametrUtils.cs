
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITreningLibrary
{
    public class ParametrUtils
    {
        /*
         * Определение выбора типа значения для записи параметров
         */

        private static void ParametrTypeVolume<T>(Parameter componentParam, T valueParam)
        {
            switch (valueParam.GetType().Name)
            {
                case "Int32":
                    componentParam.Set((int)(object)valueParam);
                    break;
                case "String":
                    componentParam.Set((string)(object)valueParam);
                    break;
                case "Double":
                    componentParam.Set((double)(object)valueParam);
                    break;
                default:
                    break;
            }
        }
        /*
         **** ==== Добавление параметра ==== ****
         */
        public bool AddParam(Application application, Document document,
            string parametrName, CategorySet categorySet, 
            BuiltInParameterGroup builtInParameterGroup,
            bool isInstance)
        {
            using (Transaction ts = new Transaction(document, "Add parametr"))
            {
                ts.Start();
                DefinitionFile definitionFile = application.OpenSharedParameterFile();
                if (definitionFile == null)
                {
                    TaskDialog.Show("Ошибка", "Файл общих параметров не найден.");
                    ts.Commit();
                    return false;
                }
                Definition definition = definitionFile.Groups
                    .SelectMany(group => group.Definitions)
                    .FirstOrDefault(def => def.Name.Equals(parametrName));
                if (definition == null)
                {
                    TaskDialog.Show("Ошибка", "Указанный параметр не найден");
                    ts.Commit();
                    return false;
                }
                Binding binding = application.Create.NewTypeBinding(categorySet);
                if (isInstance)
                {
                    binding = application.Create.NewInstanceBinding(categorySet);
                }
                BindingMap map = document.ParameterBindings;
                map.Insert(definition, binding, builtInParameterGroup);
                ts.Commit();
                return true;
            }
        }

        /*
         **** ==== Установка значения параметра в экземпляр ==== ****
         */
        public static void SetParametrFamilyInstance<T>(Document doc, FamilyInstance elem, string paramName, T valueParam)
        {
            using (Transaction tsParam = new Transaction(doc, "Set parametrs"))
            {
                tsParam.Start();
                Parameter componentParam = elem.LookupParameter(paramName);
                ParametrTypeVolume(componentParam, valueParam);
                tsParam.Commit();
            }
        }

        public static void SetParametrFamilyInstance<T>(Document doc, FamilyInstance elem, BuiltInParameter paramName, T valueParam)
        {
            using (Transaction tsParam = new Transaction(doc, "Set parametrs"))
            {
                tsParam.Start();
                Parameter componentParam = elem.get_Parameter(paramName);
                ParametrTypeVolume(componentParam, valueParam);
                tsParam.Commit();
            }
        }

        /*
         **** ==== Установка значения параметра на лист ==== ****
         */
        public static void SetParametrViewSheet<T>(Document doc, ViewSheet elem, string paramName, T valueParam)
        {
            using (Transaction tsParam = new Transaction(doc, "Set parametrs"))
            {
                tsParam.Start();
                Parameter componentParam = elem.LookupParameter(paramName);
                ParametrTypeVolume(componentParam, valueParam);
                tsParam.Commit();
            }
        }

        public static void SetParametrViewSheet<T>(Document doc, ViewSheet elem, BuiltInParameter paramName, T valueParam)
        {
            using (Transaction tsParam = new Transaction(doc, "Set parametrs"))
            {
                tsParam.Start();
                Parameter componentParam = elem.get_Parameter(paramName);
                ParametrTypeVolume(componentParam, valueParam);
                tsParam.Commit();
            }
        }
    }
}
