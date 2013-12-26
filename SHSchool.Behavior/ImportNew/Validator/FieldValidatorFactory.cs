using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;

namespace SHSchool.Behavior
{
    class FieldValidatorFactory : IFieldValidatorFactory
    {
        #region IFieldValidatorFactory 成員

        public IFieldValidator CreateFieldValidator(string typeName, System.Xml.XmlElement validatorDescription)
        {
            switch (typeName.ToUpper())
            {
                case "STUDENTNUMBEREXISTENCE":
                    return new StudentNumberExistenceValidator();
                case "STUDENTNUMBERREPEAT":
                    return new StudentNumberRepeatValidator();
                case "COMMENTRESOLVE":
                    return new CommentValidator(validatorDescription);
                default:
                    return null;
            }
        }

        #endregion
    }
}
