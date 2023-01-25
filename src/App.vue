<template>
  <div id="app" class="bg-gray-200 h-screen flex flex-col justify-start items-center">
    <!-- Header -->
    <div class="w-full p-4 bg-blue-400 flex justify-center items-center">
      <h1 class="text-xl text-gray-50">
        Autorisation de conduite
      </h1>
    </div>
    <!-- general informations --> 
    <div class="bg-gray-50 p-4 max-[200px]">
      <div>
        <h2 class="text-md font-normal">
          Informations genérales:
        </h2>
        <!-- limit date input -->
        <div class="flex flex-col mt-4">
          <label for="limitDate" class="block mb-2 text-sm font-extralight">Date d'expiration (en jrs) :</label>
          <input id="limitDate" type="number" min="0" placeholder="0" v-model="limitDate" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-1.5" required/>
        </div>
        <!-- date authorization input -->
        <div class="flex flex-col mt-4">
          <label for="dateAuthorization" class="block mb-2 text-sm font-extralight">Date du document :</label>
          <input id="dateAuthorization" type="date" v-model="dateAuthorization" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-1.5" required/>
        </div>
      </div>
      <!-- select user row input -->
      <div class="my-2">
        <h2 class="text-md font-normal">
          Conducteur/trice d'engins :
        </h2>
        <div class="flex flex-col mt-4">
          <label for="row" class="block mb-2 text-sm font-extralight">Entrer le numéro de ligne de la personne :</label>
          <input id="row" type="number" min="3" placeholder="3" v-model="rowInput" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-1.5" @change="readData()" />
        </div>
      </div>
      <!-- user selected and button for edit document -->
      <div v-if="userDataChecked" class="rounded-md border border-blue-500 flex flex-col items-center mt-4 p-4">
        <h3 class="w-full text-center text-md">Utilisateur sélectionné :</h3>
        <!-- display the user selected and button "edit document" -->
        <div v-if="userDataInfos.isPresent && medicalExaminationChecked" class="w-full flex flex-col items-center">
          <p class="w-full text-center text-blue-600 mt-2">
            {{ userDataInfos.lastName }} {{ userDataInfos.firstName }}
          </p>
          <button v-if="userDataInfos.isPresent" class="bg-blue-600 text-gray-200 py-2 px-3 rounded-md transition duration-500 hover:bg-blue-800 hover:text-gray-50 mt-4" @click="editDocument">
            Générer le document
          </button>
        </div>
        <!-- error message -->
        <div v-else>
          <p v-if="!userDataInfos.isPresent" class="text-sm text-red-500">
            {{ userDataInfos.lastName }} {{ userDataInfos.firstName }} ne fait plus parti des effectifs.
          </p>
          <p v-else-if="!medicalExaminationChecked" class="text-sm text-red-500">
            {{ userDataInfos.lastName }} {{ userDataInfos.firstName }} n'a pas de date de visite médicale valide.
          </p>
        </div>
      </div>
      <div v-else-if="rowInput > 3" class="rounded-md border border-red-500 mt-4 p-4">
        <p class="text-red-500 font-medium">
          Aucune valeur pour cette ligne
        </p>
      </div>
    </div>
  </div>
</template>

<script>
import moment from 'moment-msdate';

export default {
  name: 'App',
  data() {
    return {
      limitDate: 7,
      rowInput: 3,
      dateAuthorization: Date.now(),
      keyDataInfos : ["isPresent", "team", "lastName", "firstName", "position", "company" ],
      keyDataDates: [ "workAtHeight", "AIPR", "shotCrete", "scaffoldingReception", "R482CatA", "R482CatB1", "R482CatB2", "R482CatB3", "R482CatC2", "R482CatC1", "R3725", "R482CatC3", "R482CatD", "R482CatE", "R482CatF", "R482CatG", "OptionTMS", "OptionWinch", "R486CatA1A", "R486CatB1B", "R486CatA3A", "R486CatB3B", "R483CatB1B", "R483CatA1A", "R483CatA2A", "R483CatB2B", "R484Cat1", "R484Cat2", "R487Cat1", "R487Cat2", "R487Cat3", "R489Cat1A", "R489Cat1B", "R489Cat2A", "R489Cat2B", "R489Cat3", "R489Cat4", "R489Cat5", "R489Cat6", "R489Cat7", "R485Cat1", "R485Cat2", "R490" ],
      medicalExamination: null,
      medicalExaminationChecked: false,
      userDataInfos: {},
      userDataDates: {},
      userDataDatesChecked: {}
    }
  },
  computed: {
    userDataChecked() {
      if( this.rowInput >= 3 && this.userDataInfos && (this.userDataInfos.lastName && this.userDataInfos.lastName !== "") &&  (this.userDataInfos.lastName && this.userDataInfos.lastName !== "")) {
        return true;
      }
      return false;
    }
  },
  methods: {
    // we get the user data into cells
    async readData() {
      // if the value is less than 3, we stop processing
      if(this.rowInput.value < 3) {
        return;
      }
      // we define the different addressRange
      const addressMedicalExamination = `G${this.rowInput}:G${this.rowInput}`;
      const addressDataInfos = `A${this.rowInput}:F${this.rowInput}`;
      const addressDataDates = `I${this.rowInput}:AZ${this.rowInput}`;

      // we call functions
      await this.getData("infos", addressDataInfos)
      // we check if the user is present on the site, otherwise we stop proccessing
      if(this.userDataInfos.isPresent) {
        await this.getData("medical", addressMedicalExamination);
        await this.getData("dates", addressDataDates);
      } else {
        return;
      }
    },
    // we get data from excel cells
    async getData(type, addressDataInfos) {
      return new Promise((resolve) => {
        window.Excel.run(async (context) => {
          const ws = context.workbook.worksheets.getItem("CACES");
          const range = ws.getRange(addressDataInfos);
          range.load("values");
          await context.sync();

          if(type === "medical") {
            this.medicalExamination = range.values[0][0];
            this.transformDateExcel("medical");
          } else {
            this.initObject(type, range.values);
          }
          resolve(context);
        });
      })
    },
    // we init object : userDataInfos && userDataDates
    initObject(type, rangeValues) {

      if(type === "infos") {
        this.userDataInfos = {};

        for (let i = 0; i < this.keyDataInfos.length; i++) {
          this.userDataInfos[this.keyDataInfos[i]] = rangeValues[0][i];
        }

        this.userDataInfos.isPresent = this.userDataInfos.isPresent === "Présent";

      } else if (type === "dates") {
        this.userDataDates = {};

        for (let i = 0; i < this.keyDataDates.length; i++) {
          this.userDataDates[this.keyDataDates[i]] = rangeValues[0][i];
        }

        this.transformDateExcel("dates");
      } else {
        console.log('error type in initObject');
      }
    },
    // we transform the date excel into moment date
    transformDateExcel(type) {
      if(type === "dates") {
        Object.keys(this.userDataDates).forEach(item => {
  
          const userDataDatesValue = this.userDataDates[item];
  
          if(userDataDatesValue !== undefined && userDataDatesValue !== null && userDataDatesValue !== "") {
            this.userDataDates[item] = moment.fromOADate(this.userDataDates[item]);
          }
        })
        this.isTheDateValid(type);

      } else if(type === "medical") {
        if(this.medicalExamination !== "") {
          this.medicalExamination = moment.fromOADate(this.medicalExamination);
          this.isTheDateValid(type);
        } else {
          this.medicalExamination = null;
          this.medicalExaminationChecked = false;
        }
      }
    },
    // we check if the examination dates are valid
    isTheDateValid(type) {
      if(type === "dates") {
        this.userDataDatesChecked = {};
  
        Object.keys(this.userDataDates).forEach(item => {
  
          const userDataDatesValue = this.userDataDates[item];
  
          if(userDataDatesValue !== undefined && userDataDatesValue !== null && userDataDatesValue !== "") {

            this.userDataDatesChecked[item] = this.diffDate(this.userDataDates[item]);
          } else {
            this.userDataDatesChecked[item] = false;
          }
        })
      } else if(type === "medical") {
        this.medicalExaminationChecked = this.diffDate(this.medicalExamination);
      }
    },
    // we calculate the difference between a date and today
    diffDate(momentDateCert, momentDateNow = moment(Date.now())) {

      const diffDate = momentDateCert.diff(momentDateNow, 'days');

      if(diffDate >= this.limitDate) {
        return true;
      } else {
        return false;
      }
    },
    editDocument() {
      const minExamDate = this.calculateMinDate();
      this.writeUserDataInfos(minExamDate);
      this.writeUserDataDatesSimple();
      const categoryNames = [ "R486", "R483", "R487", "R485" ];

      categoryNames.forEach(categoryName => {
        this.prepareUserDataDatesComplex(categoryName);
      })
      console.log(this.userDataDatesChecked);
      console.log(this.userDataDates);
    },
    calculateMinDate() {
      const arrayValidDate = [];
      Object.keys(this.userDataDates).forEach(item => {
  
        if(typeof this.userDataDates[item] === 'object') {
          if(this.diffDate(this.userDataDates[item])) {
            arrayValidDate.push(moment(this.userDataDates[item]))
          }
        }

      })
      return moment.min(arrayValidDate);
    },
    writeUserDataInfos(minExamDate) {
      window.Excel.run((context) => {
        const ws = context.workbook.worksheets.getItem('TEMPLATE');
        // write lastName
        const rangeLastName = ws.getRange("D3");
        rangeLastName.values = [[this.userDataInfos.lastName]];
        // write firstName
        const rangeFirstName = ws.getRange("L3");
        rangeFirstName.values = [[this.userDataInfos.firstName]];
        // write authorization date
        const dateAuthorizationFormat = moment(this.dateAuthorization).toOADate();
        // S5
        const rangeDateAuthorization = ws.getRange("S3");
        rangeDateAuthorization.values = [[dateAuthorizationFormat]];
        rangeDateAuthorization.numberFormat = [["dd/[$-409]mm/yyyy"]];
        // E58
        const rangeDateDelivered = ws.getRange("E56");
        rangeDateDelivered.values = [[dateAuthorizationFormat]];
        rangeDateDelivered.numberFormat = [["dd/[$-409]mm/yyyy"]];
        // write medical examination
        const dateExaminationMedicalFormat = moment(this.medicalExamination).toOADate();
        const rangeDateExaminationMedical = ws.getRange("I6");
        rangeDateExaminationMedical.values = [[dateExaminationMedicalFormat]];
        rangeDateExaminationMedical.numberFormat = [["dd/[$-409]mm/yyyy"]];
        // write expiration date
        const dateExpirationFormat = moment(minExamDate).toOADate();
        const rangeDateExpiration = ws.getRange("S6");
        rangeDateExpiration.values = [[dateExpirationFormat]];
        rangeDateExpiration.numberFormat = [["dd/[$-409]mm/yyyy"]];
        // write lastname + firstName
        const rangeFullName = ws.getRange("D26");
        rangeFullName.values = [[this.userDataInfos.lastName + " " + this.userDataInfos.firstName]];
        return context.sync();
      });
    },
    writeUserDataDatesSimple() {
      window.Excel.run((context) => {
        const ws = context.workbook.worksheets.getItem('TEMPLATE');
        // write R482 Cat A
        const rangeR482CatA = ws.getRange("B28");
        rangeR482CatA.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R482CatA)]];
        // write R482 Cat B1
        const rangeR482CatB1 = ws.getRange("B29");
        rangeR482CatB1.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R482CatB1)]];
        // write R482 Cat B2
        const rangeR482CatB2 = ws.getRange("B31");
        rangeR482CatB2.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R482CatB2)]];
        // write R482 Cat C1
        const rangeR482CatC1 = ws.getRange("B32");
        rangeR482CatC1.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R482CatC1)]];
        // write R482 Cat C2
        const rangeR482CatC2 = ws.getRange("B33");
        rangeR482CatC2.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R482CatC2)]];
        // write R482 Cat C3
        const rangeR482CatC3 = ws.getRange("B34");
        rangeR482CatC3.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R482CatC3)]];
        // write R482 Cat D
        const rangeR482CatD = ws.getRange("B35");
        rangeR482CatD.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R482CatD)]];
        // write R482 Cat E
        const rangeR482CatE = ws.getRange("B36");
        rangeR482CatE.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R482CatE)]];
        // write R482 Cat F
        const rangeR482CatF = ws.getRange("B37");
        rangeR482CatF.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R482CatF)]];
        // write option winch (F)
        const rangeOptionWinch = ws.getRange("C38");
        rangeOptionWinch.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.OptionWinch)]];
        // write R482 Cat G
        const rangeR482CatG = ws.getRange("B39");
        rangeR482CatG.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R482CatG)]];
        // write R486 Cat A 1A
        const rangeR486CatA1A = ws.getRange("E41");
        rangeR486CatA1A.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R486CatA1A)]];
        // write R486 Cat A 3A
        const rangeR486CatA3A = ws.getRange("G41");
        rangeR486CatA3A.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R486CatA3A)]];
        // write R486 Cat B 1B
        const rangeR486CatB1B = ws.getRange("E42");
        rangeR486CatB1B.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R486CatB1B)]];
        // write R486 Cat B 3B
        const rangeR486CatB3B = ws.getRange("G42");
        rangeR486CatB3B.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R486CatB3B)]];
        // write R483 Cat A 1A
        const rangeR483CatA1A = ws.getRange("H44");
        rangeR483CatA1A.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R483CatA1A)]];
        // write R483 Cat A 1B
        const rangeR483CatB1B = ws.getRange("J44");
        rangeR483CatB1B.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R483CatB1B)]];
        // write R483 Cat A 2A
        const rangeR483CatA2A = ws.getRange("H45");
        rangeR483CatA2A.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R483CatA2A)]];
        // write R483 Cat A 2B
        const rangeR483CatB2B = ws.getRange("J45");
        rangeR483CatB2B.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R483CatB2B)]];
        // write R489 Cat 1A
        const rangeR489Cat1A = ws.getRange("L28");
        rangeR489Cat1A.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R489Cat1A)]];
        // write R489 Cat 1B
        const rangeR489Cat1B = ws.getRange("L29");
        rangeR489Cat1B.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R489Cat1B)]];
        // write R489 Cat 3
        const rangeR489Cat3 = ws.getRange("L30");
        rangeR489Cat3.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R489Cat3)]];
        // write R489 Cat 4
        const rangeR489Cat4 = ws.getRange("L31");
        rangeR489Cat4.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R489Cat4)]];
        // write R489 Cat 5
        const rangeR489Cat5 = ws.getRange("L32");
        rangeR489Cat5.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R489Cat5)]];
        // write R489 Cat 6
        const rangeR489Cat6 = ws.getRange("L33");
        rangeR489Cat6.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R489Cat6)]];
        // write R489 Cat 7
        const rangeR489Cat7 = ws.getRange("L34");
        rangeR489Cat7.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R489Cat7)]];
        // write R484 Cat 1
        const rangeR484Cat1 = ws.getRange("L35");
        rangeR484Cat1.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R484Cat1)]];
        // write R485 Cat 1
        const rangeR485Cat1 = ws.getRange("M38");
        rangeR485Cat1.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R485Cat1)]];
        // write R485 Cat 2
        const rangeR485Cat2 = ws.getRange("M39");
        rangeR485Cat2.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R485Cat2)]];
        // write R487 Cat 1
        const rangeR487Cat1 = ws.getRange("M41");
        rangeR487Cat1.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R487Cat)]];
        // write R487 Cat 2
        const rangeR487Cat2 = ws.getRange("M42");
        rangeR487Cat2.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R487Cat2)]];
        // write R487 Cat 3
        const rangeR487Cat3 = ws.getRange("M43");
        rangeR487Cat3.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R487Cat3)]];
        // write R490
        const rangeR490 = ws.getRange("L44");
        rangeR490.values = [[this.transformBooleanInNumber(this.userDataDatesChecked.R490)]];
        return context.sync();
      });
    },
    prepareUserDataDatesComplex(categoryName) {

        if(categoryName === "R486") {
          // we define the arrays with different values for filter
          const categoryA = [ "R486CatA1A", "R486CatA3A" ];
          const categoryB = [ "R486CatB1B", "R486CatB3B" ];
          const category = categoryA.concat(categoryB);
          // we define if at least one is checked
          const isCategoryChecked = category.some(item => this.userDataDatesChecked[item])
          const isCategoryAChecked = categoryA.some(item => this.userDataDatesChecked[item])
          const isCategoryBChecked = categoryB.some(item => this.userDataDatesChecked[item])
          // we call write function with address and value
          this.writeUserDataDatesLoop("B40", isCategoryChecked)
          this.writeUserDataDatesLoop("C41", isCategoryAChecked)
          this.writeUserDataDatesLoop("C42", isCategoryBChecked)
        } else if(categoryName === "R483") {
          // we define the arrays with different values for filter
          const categoryA = [ "R483CatA1A", "R483CatA2A" ];
          const categoryB = [ "R483CatB1B", "R483CatB2B" ];
          const category = categoryA.concat(categoryB);
          // we define if at least one is checked
          const isCategoryChecked = category.some(item => this.userDataDatesChecked[item])
          const isCategoryAChecked = categoryA.some(item => this.userDataDatesChecked[item])
          const isCategoryBChecked = categoryB.some(item => this.userDataDatesChecked[item])
          // we call write function with address and value
          this.writeUserDataDatesLoop("B43", isCategoryChecked)
          this.writeUserDataDatesLoop("C44", isCategoryAChecked)
          this.writeUserDataDatesLoop("C45", isCategoryBChecked)
        } else if(categoryName === "R485") {
          // we define the arrays with different values for filter
          const category = [ "R485Cat1", "R485Cat2" ];
          // we define if at least one is checked
          const isCategoryChecked = category.some(item => this.userDataDatesChecked[item])
          // we call write function with address and value
          this.writeUserDataDatesLoop("L37", isCategoryChecked)
        } else if(categoryName === "R487") {
          // we define the arrays with different values for filter
          const category = [ "R487Cat1", "R487Cat2", "R487Cat3" ];
          // we define if at least one is checked
          const isCategoryChecked = category.some(item => this.userDataDatesChecked[item])
          // we call write function with address and value
          this.writeUserDataDatesLoop("L40", isCategoryChecked)
        } else {
          console.log('error type category in prepareUserDataDatesComplex')
        }

    },
    writeUserDataDatesLoop(address, value) {
      window.Excel.run((context) => {
        const ws = context.workbook.worksheets.getItem('TEMPLATE');
        
        // write
        const rangeOptionWinch = ws.getRange(address);
        rangeOptionWinch.values = [[this.transformBooleanInNumber(value)]];

        return context.sync();
      })
    },
    transformBooleanInNumber(value) {
      return value === true ? 1 : 0;
    }
  }
};
</script>