using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParametricBuilder.Models
{
    public interface IFileHandler
    {
        void HandleFileSelected(string filePath);
    }
}
