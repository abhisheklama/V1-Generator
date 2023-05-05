const xlsx = require("xlsx");
const planSheet = xlsx.readFile("./Input/william/benefits.xlsx");
const rate = xlsx.readFile("./Input/william/rateSheet.xlsx");
let GlobalData = xlsx.utils.sheet_to_json(
  planSheet.Sheets[planSheet.SheetNames[0]]
);
let rateSheet = xlsx.utils.sheet_to_json(rate.Sheets[rate.SheetNames[0]]);
let conversion = 3.6725;
let fs = require("fs");
const jsonToCSV = require("json-to-csv");

function createSheet() {
  try {
    let annual = rateSheet.filter((v) => v.frequency == "Annually");
    let arr = annual.map((rate) => {
      // let quater = rateSheet.find(
      //   (v) =>
      //     v.frequency == "Quarterly" &&
      //     rate.ageStart == v.ageStart &&
      //     rate.gender == v.gender &&
      //     rate.copay == v.copay &&
      //     rate.planName == v.planName &&
      //     rate.network == v.network &&
      //     rate.coverage == v.coverage
      // );
      let benefits = GlobalData.find((v) => {
        if (v["Plan Name"] == "All") return true;
        return v["Plan Name"] == rate.planName;
      });
      let struc = {
        PlanName1: rate.planName,
        PlanName2: rate.network,
        rateMonth: parseFloat(rate.monthly) / conversion,
        rateQuarter: parseFloat(rate.quaterly) / conversion,
        rateSemiAnnual: parseFloat(rate.semi) / conversion,
        rateAnnual: parseFloat(rate.rates) / conversion,
        rateBiannual: "",
        ageRangeStart: rate.ageStart,
        ageRangeEnd: rate.ageEnd,
        Gender: rate.gender,
        Currency: "USD",
        insuranceCoverAmount: "",
        InsurerArea: rate.coverage,
        InsurerAreaEx: "",
        InsuranceArea: "",
        InsuranceAreaEx: "",
        TripDuration: "",
        Child: "",
        physicianDeductible: "",
        physicianDeductibleMax: "",
        coPay: rate.copay,
        coPayOn: "",
        deductable: "",
        Terms1: "",
        Terms2: "",
        Terms3: "",
        Terms4: "",
        Terms5: "",
        Terms6: "",
        Terms7: "",
        annualDentalPrimary: "",
        semiAnnualDentalPrimary: "",
        quarterlyDentalPrimary: "",
        annualDentalPrimarySpouse: "",
        semiAnnualDentalPrimarySpouse: "",
        quarterlyDentalPrimarySpouse: "",
        annualDentalPrimaryChildren: "",
        semiAnnualDentalPrimaryChildren: "",
        quarterlyDentalPrimaryChildren: "",
        annualDentalPrimarySpouseChildren: "",
        semiAnnualDentalPrimarySpouseChildren: "",
        quarterlyDentalPrimarySpouseChildren: "",
        addon1: "",
        addon2: "",
        addon3: "",
        addon4: "",
        addon5: "",
        addon6: "",
        addon7: "",
        addon8: "",
        addon9: "",
        addon10: "",
        addon11: "",
        addon12: "",
        addon13: "",
        addon14: "",
        addon15: "",
        addon16: "",
        addon17: "",
        addon18: "",
        AnnualLimit:
          benefits["Annual Limit"] % 2 == 0
            ? "AED " + benefits["Annual Limit"]
            : benefits["Annual Limit"],
        InPatientDirectBilling: "",
        OutPatientDirectBilling: "",
        OutOfNetworkClaimsHandling: benefits["Claims Handling"],
        InpatientHospitalisation:
          benefits["In-patient (Hospitalization & Surgery)"],
        OutPatient: benefits["Out-patient benefits"],
        Physiotherapy: benefits["Physiotherapy"],
        EmergencyEvacuation: benefits["Emergency Evacuation"],
        ChronicConditions: benefits["Chronic Condition Cover"],
        PreExistingCover: benefits["Pre-existing Condition Cover"],
        RoutineMaternity:
          benefits["Maternity (Consultations, Scans and Delivery)"],
        MaternityWaitingPeriod: benefits["Maternity Waiting Period"],
        ComplicationsOfPregnancy: benefits["Complications of Pregnancy"],
        NewBornCoverage: benefits["New Born Cover"],
        Dental: benefits["Dental"],
        DentalWaitingPeriod: benefits["Dental Waiting Period"],
        OpticalBenefits: benefits["Optical Benefits"],
        Wellness: benefits["Wellness & Health Screening"],
        SemiAnnualSurcharge: benefits["Semi Annual Surcharge"],
        QuarterlySurcharge: benefits["Quarterly Surcharge"],
        MonthlySurcharge: benefits["Monthly Surcharge"],
        RoutineMaternityFilter: benefits["Routine Maternity"].toLowerCase(),
        WellnessFilter: benefits["Wellness"].toLowerCase(),
        OpticalFilter: benefits["Optical"].toLowerCase(),
        CompanyName: GlobalData[0]["companyName"],
        StartDate: GlobalData[0]["startDate"],
        EndDate: GlobalData[0]["endDate"],
        Residency: GlobalData[0]["residency"],
        a1: "",
        a2: "",
        dentalFilter: 2,
        // singleFemale:
        //   rate.married == 1 || rate.married == 0 ? rate.married : "",
      };
      if (struc.OutPatient.includes("$")) {
        let value = GlobalData.find((v) => v.copay.split("/")[0] == rate.copay);
        let copay = value.copay.split("/");
        copay.forEach((v, index) => {
          if (index == 0) return;
          struc.OutPatient = struc.OutPatient.replace("$", v);
        });
        // if (rate.copay == "Nil") {
        //   struc.OutPatient =
        //     "Medicines and diagnostics & lab tests covered in full Consultations covered with Nil co-pay or 20% co-pay";
        // }
      }
      for (let key in struc) {
        while (typeof struc[key] == "string" && struc[key].includes("\n")) {
          struc[key] = struc[key].replace("\n", " ");
        }
      }
      // if (struc.PlanName1 == "Gold") {
      // struc.Dental = "Dental Plus- Covered up to AED 7,340 with 20% co-pay";
      // struc.DentalWaitingPeriod = "Dental Plus- 12 months wait";
      // struc.Dental = "Optional";
      // struc.DentalWaitingPeriod = "No wait";
      // } else {
      // struc.Dental = "Dental Plus- Covered up to AED 5,505 with 20% co-pay";
      // struc.DentalWaitingPeriod = "Dental Plus- 12 months wait";
      // struc.Dental = "Optional";
      // struc.DentalWaitingPeriod = "No wait";
      // }

      return struc;
    });
    // console.log("count-", arr.length);
    arr = arr.filter((v) => v.PlanName1 == "Silver");
    let newArr = [];

    for (i = 1; i <= 3; i++) {
      newArr.push([...arr.splice(0, 800)]);
    }

    newArr.forEach((v, i) => {
      jsonToCSV(v, `Output/silver-${i}.csv`)
        .then(() => {
          console.log("Sheet Generated Successfully!");
        })
        .catch((error) => {
          console.log("Something went wrong");
          console.log({ err: error });
        });
    });

    // jsonToCSV(arr, `Output/${GlobalData[0].companyName}.csv`)
    //   .then(() => {
    //     console.log("Sheet Generated Successfully!");
    //   })
    //   .catch((error) => {
    //     console.log("Something went wrong");
    //     console.log({ err: error });
    //   });
  } catch (error) {
    console.log(error);
    console.log({ err: error.message, stack: error.stack });
  }
}

createSheet();
