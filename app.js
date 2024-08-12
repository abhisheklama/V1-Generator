const xlsx = require("xlsx");
const planSheet = xlsx.readFile("./Input/Medgulf_Medex/benefit.xlsx");
let GlobalData = xlsx.utils.sheet_to_json(
  planSheet.Sheets[planSheet.SheetNames[0]]
);
const rate = xlsx.readFile("./Input/Medgulf_Medex/rateSheet.xlsx");
let rateSheet = xlsx.utils.sheet_to_json(rate.Sheets[rate.SheetNames[0]]);
let conversion = 3.6725;
let fs = require("fs");
const jsonToCSV = require("json-to-csv");

function createSheet() {
  try {
    let annual = rateSheet.filter((v) => v.frequency == "Annually");
    let arr = annual.map((rate, i) => {
      // if (rate.copay == "0.1") rate.copay = "10%";
      // if (rate.copay == "0.2") rate.copay = "20%";

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
      if (!benefits) {
        console.log(rate.planName);
        throw new Error("plan not found");
      }
      for (let key in benefits) {
        if (
          benefits[key] &&
          benefits[key].toString().toLowerCase().includes("nishima")
        )
          throw new Error("Nishma ALERT!!!!!!!!!!!!!!!!");
      }

      let struc = {
        PlanName1: rate.planName,
        PlanName2: rate.network,
        rateMonth: "", //rate.month,
        // parseFloat(rate.monthly) / conversion +
        // parseFloat(rate.dental / 12) / conversion,
        rateQuarter: "", //parseFloat(quater.rates) / conversion,
        // parseFloat(rate.quaterly) / conversion +
        // parseFloat(rate.dental / 4) / conversion,
        rateSemiAnnual: "",
        // parseFloat(rate.semi) / conversion +
        // parseFloat(rate.dental / 2) / conversion,
        rateAnnual: parseFloat(rate.rates) / conversion,
        // parseFloat(rate.dental) / conversion,
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
        Terms4: "", //rate.dental,
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
        OpticalFilter: benefits["Optical"]?.toLowerCase(),
        CompanyName: GlobalData[0]["companyName"],
        StartDate: GlobalData[0]["startDate"],
        EndDate: "",
        CJ: "",
        CK: "",
        CL: rate?.type ? rate.type : "",
        CM: benefits["Dental Filter"]?.toLowerCase() == "yes" || true ? 0 : "",
        CN: "",
        CO: "",
        CP: "",
        CQ: "",
        CR: "",
        CS: "",
        CT: "",
        CU: "",
        CV: "",
        CW: "",
        CX: "",
        CY: "",
        CZ: "",
        DA: "",
        AccommodationType: benefits["Accommodation Type"],
        DC: GlobalData[0]["endDate"],
        Residency: rate.residency ?? GlobalData[0]["residency"],
        relation: rate.relation ? rate.relation : "",
        singleFemale:
          rate.married == 1 || rate.married == -1 || rate.married == 0
            ? rate.married
            : "",
        singleChild: "",
        dentalAddon:
          rate.dental && false ? parseFloat(rate.dental) / conversion : "",
      };
      // struc.Dental =
      //   "Routine dental- Covered up to USD 250 with 20% co-pay Complex dental- Covered up to USD 1,000 with 20% co-pay";
      // struc.DentalWaitingPeriod = "9 months wait";
      if (GlobalData[0].comment) {
        struc.OutPatient = struc.OutPatient.replace(GlobalData[0].comment, "");
      }
      if (struc.OutPatient.includes("$")) {
        let value = GlobalData.find(
          (v) => v?.copay?.split("/")[0] == rate.copay
        );
        if (!value) console.log("rate.copay --> ", rate.copay, i);
        let copay = value.copay.split("/");
        copay.forEach((v, index) => {
          ``;
          if (index == 0) return;
          struc.OutPatient = struc.OutPatient.replace("$", v);
        });
        // if (rate.copay == "Nil") {
        //   struc.OutPatient =
        //     "Medicines and diagnostics & lab tests covered in full Consultations covered with Nil co-pay or 20% co-pay";
        // }
      }
      if (struc.Physiotherapy.includes("$")) {
        let value = GlobalData.find((v) => v.copay.split("/")[0] == rate.copay);
        let copay = value.copay.split("/");
        // copay.forEach((v, index) => {
        //   if (index == 0) return;
        struc.Physiotherapy = struc.Physiotherapy.replace("$", copay[1]);
        // });
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
      //   struc.Dental = "Dental Plus- Covered up to AED 7,340 with 20% co-pay";
      //   struc.DentalWaitingPeriod = "Dental Plus- 12 months wait";
      //   struc.Dental = "Optional";
      //   struc.DentalWaitingPeriod = "No wait";
      // } else {
      //   struc.Dental = "Dental Plus- Covered up to AED 5,505 with 20% co-pay";
      //   struc.DentalWaitingPeriod = "Dental Plus- 12 months wait";
      //   struc.Dental = "Optional";
      //   struc.DentalWaitingPeriod = "No wait";
      // }

      return struc;
    });

    console.log("count-", arr.length);
    let newArr = [];
    let len = Math.ceil(arr.length / 800);

    if (arr.length > 800) {
      for (i = 1; i <= len; i++) {
        if (i == len) {
          newArr.push([...arr.splice(0, 800)]);
          newArr.push([...arr.splice(0, arr.length - 1)]);
        } else newArr.push([...arr.splice(0, 800)]);
      }
    }
    console.log("new-", newArr.length);

    if (newArr.length != 0) {
      newArr.forEach((v, i) => {
        jsonToCSV(
          v,
          `Output/${GlobalData[0].companyName}-${i}-${
            GlobalData[0].residency.includes("- ")
              ? GlobalData[0].residency.split("- ")[1]
              : GlobalData[0].residency
          }.csv`
        )
          .then(() => {
            console.log("Sheet Generated Successfully!");
          })
          .catch((error) => {
            console.log("Something went wrong");
            console.log({ err: error });
          });
      });
    } else {
      jsonToCSV(
        arr,
        `Output/${GlobalData[0].companyName}-${
          GlobalData[0].residency.includes("- ")
            ? GlobalData[0].residency.split("- ")[1]
            : GlobalData[0].residency
        }.csv`
      )
        .then(() => {
          console.log("Sheet Generated Successfully!");
        })
        .catch((error) => {
          console.log("Something went wrong");
          console.log({ err: error });
        });
    }
  } catch (error) {
    console.log(error);
    console.log({ err: error.message, stack: error.stack });
  }
}

createSheet();
