using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;

using Alma.BaseUI.Utils;

using Wpm.Schema.Kernel;
using Wpm.Implement.Manager;
using Wpm.Implement.ComponentEditor;


using Actcut.QuoteModel;
using Actcut.QuoteModelManager;
using Actcut.ActcutModelManagerUI;

namespace Actcut.ActcutClipperApi
{
    public class ClipperApi : IClipperApi
    {
        IContext _Context = null;
        bool _UserOk = false;

        #region IClipperApi Membres

        public bool Init(IContext context)
        {
            _Context = context;
            _UserOk = (_Context.UserId > 0);
            return true;
        }

        public bool Init(string databaseName, string user)
        {
            ModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(databaseName);

            if (_Context != null)
            {
                Licence.InitLicence(_Context.Kernel, null);
                _UserOk = SetUser(_Context, user);
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool ExportQuote(long quoteNumber, string orderNumber, string exportFile)
        {
            if (_Context != null)
            {
                IQuoteManagerUI quoteManagerUI = new QuoteManagerUI();
                IEntity quoteEntity = quoteManagerUI.GetQuoteEntity(_Context, quoteNumber);
                
                ////////// temporaire//////////////////////////////////////////
                ///itegration api vince
                ///
                string sep = "\r\n";
                StreamWriter w = new StreamWriter(@exportFile);
                string content = "du devis 10690" +sep+
                    "IDDEVIS¤10690¤10690¤¤SIFAB¤ALMA¤¤¤¤38000¤GRENOBLE¤20181130¤20181130¤¤1¤¤1¤SUPER¤¤¤¤¤SUPER¤¤¤¤20181130¤20181130¤¤¤¤¤0¤¤9¤¤¤20181215¤¤" + sep +
                    "OFFRE¤PAPA¤10690¤1¤0.1100¤0.1100¤0.1100¤0.1100¤1¤0¤1¤1¤0¤0¤L45FM¤2¤0¤1¤+" +sep+
                    @"ENDEVIS¤10690¤SIFAB¤PAPA¤¤Forme 100 x 60¤1050 * H24 1 mm¤¤001¤PAPA¤Forme 100 x 60¤4598¤001¤3¤¤0¤1¤1¤¤¤¤¤0.0135¤0.0183¤-4598¤¤¤¤0¤0¤0¤0¤0¤0¤0¤0¤0¤0¤C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\EMF\_QUOTE_PART__PREVIEW_4598.emf¤¤0¤¤" + sep +
                    "GADEVIS¤001¤¤10¤ACHAT NOMENCLATURE¤¤¤¤¤¤NOMEN¤0¤0¤0¤0¤20181130¤10690¤¤0¤0¤0¤0¤0¤0¤0¤¤¤¤¤¤¤¤¤¤" + sep +
                    "GADEVIS¤001¤¤20¤COUPE¤¤¤¤¤¤DECOU¤0.0000¤0.0005¤0.0432¤90.0000¤20181130¤10690¤¤0¤0¤0¤0¤0¤0¤0¤¤¤¤¤¤¤¤¤¤" + sep +
                    "NOMENDVALMA¤10690¤PAPA¤0¤001¤10¤TL * 1050 * H24 * 1¤5018.2523¤0.0604¤" + sep +
                    "Fin d'enregistrement OK¤" + sep +
                    "Fin du fichier OK";



                w.Write(content);
                w.Close();
                ////////// temporaire//////////////////////////////////////////
                ///
                /// 
                ///
                bool ret = quoteManagerUI.AccepQuote(_Context, quoteEntity, orderNumber, exportFile);
                return (ret ? true : false);
            }
            else
            {
                return false;
            }
        }
        public bool GetQuote(out long quoteNumberReference)
        {
            quoteNumberReference = -1;
            if (_Context != null)
            {
                IEntity quoteEntity = null;
                IEntitySelector entitySelector = new EntitySelector();
                entitySelector.Init(_Context, _Context.Kernel.GetEntityType("_QUOTE_SENT"));
                entitySelector.MultiSelect = false;
                entitySelector.ShowPropertyBox = false;
                if (entitySelector.ShowDialog() == DialogResult.OK)
                    quoteEntity = entitySelector.SelectedEntity.FirstOrDefault();

                if (_UserOk) _Context.SaveUserModel();

                if (quoteEntity != null)
                {
                    //string quoteReference =quoteEntity.GetFieldValueAsString("ID");
                    //long quoteReference = quoteEntity.Id;
                    quoteNumberReference = quoteEntity.Id;
                    //long.TryParse(quoteReference, out quoteNumberReference);
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        #endregion

        private bool SetUser(IContext context, string user)
        {
            IEntityList userEntityList = context.EntityManager.GetEntityList("SYS_USER", "USER_NAME", ConditionOperator.Equal, user);
            userEntityList.Fill(false);

            IEntity userEntity = userEntityList.FirstOrDefault();
            if (userEntity != null)
            {
                context.UserId = userEntity.Id;
                return true;
            }

            return false;
        }
    }
}
