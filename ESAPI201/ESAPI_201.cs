using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;


namespace ESAPI201
{
    public interface IAmWellTrained
    {
        void Sit();
    }
    public class Kennel
    {
        private List<IAmWellTrained> _RegisteredPets = new List<IAmWellTrained>();

        private List<Mammal> _RegisteredMammals = new List<Mammal>();
        public Kennel() { }

        public void Register(IAmWellTrained pet)
        {
            _RegisteredPets.Add(pet);
            // I have abstracted my registration to support all pets that can "sit"
        }

        public void RegisterMammal(Mammal mammal)
        {
            _RegisteredMammals.Add(mammal);
        }

        private void TakeDogsForAWalk()
        {
            foreach (Mammal mammal in _RegisteredMammals)
            {
                if (mammal is Dog)
                {
                    // Only take dogs for a walk
                }
                if (mammal is Dog && mammal is IAmWellTrained)
                {
                    // This dog gets to go to off leash park
                }
            }

        }
    }

    public class Mammal
    {
        public bool IsWarmBlooded { get; private set; } = true;

        public virtual void Pet()
        {
            // This could be any mammal (like a Tiger) so we should probably
            // include a LOT of safety precautions and insurance forms. Slow.

        }
    }

    public class Dog : Mammal, IAmWellTrained
    {
        public string Name { get; private set; }
        public bool IsExcitable { get; private set; }

        public Dog(string name, bool isExcitable)
        {
            Name = name;
            IsExcitable = isExcitable;
        }
        public void Sit()
        {
            // My choice how to implement, but I must implement Sit() to compile if I declare the interface
        }

        public override void Pet()
        {
            // This is a well trained dog, so this will be fast and easy!

        }
    }

    public class StructureSetWrapper
    {
        private StructureSet _structureSet;

        public StructureSetWrapper(StructureSet structureSet)
        {
            _structureSet = structureSet;
        }



        public Structure GetFirstStructure()
        {
            return _structureSet.Structures.First();
            // this will throw an exception if there are no structures
        }
        public Structure GetFirstStructureOrNull()
        {
            return _structureSet.Structures.FirstOrDefault();
            // this will return null if there are no structures.
        }
        public List<Structure> GetAllPTVs()
        {
            return _structureSet.Structures.Where(x => x.DicomType == "PTV").ToList();
            // returns a list of all structures that have the "PTV" DICOM type
        }
        public List<string> GetStructureIds()
        {
            return _structureSet.Structures.Select(x => x.Id).ToList();
            // returns a list of all structure Ids
        }
        public bool IsThereAPTV()
        {
            return _structureSet.Structures.Any(x => x.DicomType == "PTV");
            // returns true if any structure is a PTV
        }

        public List<string> GetPTVsIdsAboveCertainSizeByName(string PTVname, double minVolume)
        {
            return _structureSet.Structures.Where(x => x.Id.Contains(PTVname) && x.Volume > minVolume).Select(y => y.Id).ToList();
        }

        public double getMinVolume()
        {
            return _structureSet.Structures.Aggregate(double.PositiveInfinity, (curmin, x) => curmin < x.Volume ? x.Volume : curmin);

        }


        public void writeDataToCSV(List<Structure> data)
        {
            using (var writer = new StreamWriter(@"\\datapath\myNeFile.csv"))
            {
                foreach (var s in data)
                {
                    writer.WriteLine(string.Format("{0},{1},{2}", s.Id, s.Volume, s.DicomType));
                }
            }

            List<string> stringData = new List<string>();
            using (var reader = new StreamReader(@"\\datapath\myInputFile.csv"))
            {
                foreach (var readString in reader.ReadLine().Split(','))
                {
                    stringData.Add(readString); ;
                }
            }
        }

        public void writeDataToExcel(List<Structure> data)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var ExcelFile = new ExcelPackage(@"\\datapath\myExcelFile.xslx"))
            {
                var workbook = ExcelFile.Workbook;
                var sheet1 = workbook.Worksheets.Add("sheet1");
                int rowCounter = 1;
                foreach (var s in data)
                {
                    sheet1.Cells[rowCounter, 1].Value = s.Id;
                    sheet1.Cells[rowCounter, 2].Value = s.Volume;
                    sheet1.Cells[rowCounter, 3].Value = s.DicomType;
                }


            }
        }

        public enum Pets
        {
            Cat=0,
            Dog=1,
            Fish=2
        }





        public enum PetsWithDescription
        {
            // Using System.ComponentModel
            [Description("Pet cat")] Cat = 0,
            [Description("Pet dog")] Dog = 1,
            [Description("Pet fish")] Fish = 2
        }

        public void MakeDecision(Pets thisPet)
        {
                        List<Pets> listOfPets = new List<Pets>();
            switch (thisPet)
            {
                case Pets.Dog:
                    // Take for walk
                    break;
                case Pets.Fish:
                    // Clean tank
                    break;
                case Pets.Cat:
                    // Do nothing, it's a trap
                    break;
                default:
                    // Do something any pet might enjoy
                    // Add to a list of pets
                    listOfPets.Add(thisPet);
                    break;
            }
        }

        public void MakeDecision(string thisPet)
        {
            switch (thisPet)
            {
                case "Dog":
                    // Take for walk
                    break;
                case "Fish":
                    // Clean tank
                    break;
                case "Cat":
                    // Do nothing, it's a trap
                    break;
                default:
                    // Do something any pet might enjoy
                    break;
            }
        }


    }

}
