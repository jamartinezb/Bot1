using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SelectorManager
{
    public class UiSelector
    {
        private string Sistema;
        public string NombreVentana{
            get{return Sistema;}
            set{Sistema = value;}
        }

        private string vStrSelector;
        public string Selector{
            get{return vStrSelector;}
            set{vStrSelector = value;}
        }


        private Attr Propiedades;
        public Attr Attribute{
            get{return Propiedades;}
            set{Propiedades = value;}
        }

        // Creamos la clase con los atributos
        public class Attr
        {
            private string vStrTitle = "";
            public string title{
                get{return vStrTitle;}
                set{vStrTitle = value;}
            }

            private string vStrApp = "";
            public string app{
                get{return vStrApp;}
                set{vStrApp = value;}
            }

            private string vStrId = "";
            public string id{
                get{return vStrId;}
                set{vStrId = value;}
            }

            private string vStrTag = "";
            public string tag{
                get{return vStrTag;}
                set{vStrTag = value;}
            }

            private string vStrAaname = "";
            public string aaname{
                get{return vStrAaname;}
                set{vStrAaname = value;}
            }

            private string vStrClass = "";
            public string className{
                get{return vStrClass;}
                set{vStrClass = value;}
            }

            private string vStrName = "";
            public string name{
                get{return vStrName;}
                set{vStrName = value;}
            }

            private string vStrType = "";
            public string type{
                get{return vStrType;}
                set{vStrType = value;}
            }

            private string vStrhref = "";
            public string href{
                get{return vStrhref;}
                set{vStrhref = value;}
            }

            private string vStrParentClass = "";
            public string parentclassName{
                get{return vStrParentClass;}
                set{vStrParentClass = value;}
            }

            private string vStrParentId = "";
            public string parentid{
                get{return vStrParentId;}
                set{vStrParentId = value;}
            }

            private string vStrInstance = "";
            public string instance
            {
                get { return vStrInstance; }
                set { vStrInstance = value; }
            }

            private string vStrRole = "";
            public string role
            {
                get { return vStrRole; }
                set { vStrRole = value; }
            }

            private string vStrText = "";
            public string text
            {
                get { return vStrText; }
                set { vStrText = value; }
            }
        }

        public UiSelector()
        {
            Propiedades = new Attr();
        }
    }

    public class Funciones
    {
        public static Dictionary<string, UiSelector> TransformIntoUiSelector(string vStrSelector)
        {
            List<string> vVecStrSelector;
            string vStrLine;
            Dictionary<string, UiSelector> vDicResponse = new Dictionary<string, UiSelector>();

            vVecStrSelector = vStrSelector.Split(new Char[] { '>' }).ToList();
            foreach (string Line in vVecStrSelector)
            {
                vStrLine = Line.Replace("<", "").Replace(" /", "");
                if (vStrLine.Count() > 0)
                {
                    // Crear una nueva llave en el diccionario
                    vDicResponse.Add(vVecStrSelector.IndexOf(Line).ToString(), new UiSelector());
                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].NombreVentana = System.Text.RegularExpressions.Regex.Match(vStrLine, "[a-z]*").Value;
                    foreach (System.Text.RegularExpressions.Match Att in System.Text.RegularExpressions.Regex.Matches(vStrLine, @"[a-z]*\=\'[^']*\'"))
                    {
                        switch (System.Text.RegularExpressions.Regex.Match(Att.Value, "[a-z]*").Value)
                        {
                            case "title":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.title = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "app":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.app = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "id":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.id = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "tag":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.tag = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }
                            case "aaname":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.aaname = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "class":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.className = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }
                            case "cls":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.className = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "name":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.name = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "type":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.type = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "parentclass":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.parentclassName = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "parentid":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.parentid = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "href":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.href = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "instance":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.instance = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "role":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.role = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            case "text":
                                {
                                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Attribute.text = Att.Value.Remove(0, Att.Value.IndexOf("=")).Replace("'", "").Replace("=", "");
                                    break;
                                }

                            default:
                                {
                                    break;
                                }
                        }
                    }
                    vDicResponse[vVecStrSelector.IndexOf(Line).ToString()].Selector = vStrLine.Remove(0, vStrLine.IndexOf(" "));
                }
            }

            return vDicResponse;
        }

    }

}
