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
        <!-- Responsible input -->
        <div class="flex flex-col mt-4">
          <label for="responsible" class="block mb-2 text-sm font-extralight text-gray-900">Nom et Prénom du Délégant :</label>
          <input id="responsible" type="text" v-model="responsible" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-1.5" required/>
        </div>
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
      responsible: "Alexander Heim",
      limitDate: 7,
      rowInput: 3,
      dateAuthorization: Date.now(),
      keyDataInfos : ["isPresent", "team", "lastName", "firstName", "position", "company" ],
      keyDataDates: [ "workAtHeight", "AIPR", "shotCrete", "scaffoldingReception", "R482CatA", "R482CatB1", "R482CatB2", "R482CatB3"],
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
      const addressDataDates = `I${this.rowInput}:P${this.rowInput}`;

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
        // write lastName
        const ws = context.workbook.worksheets.getItem('TEMPLATE');
        const rangeLastName = ws.getRange("D5");
        rangeLastName.values = [[this.userDataInfos.lastName]];
        // write firstName
        const rangeFirstName = ws.getRange("L5");
        rangeFirstName.values = [[this.userDataInfos.firstName]];
        // write authorization date
        const dateAuthorizationFormat = moment(this.dateAuthorization).toOADate();
        // S5
        const rangeDateAuthorization = ws.getRange("S5");
        rangeDateAuthorization.values = [[dateAuthorizationFormat]];
        rangeDateAuthorization.numberFormat = [["dd/[$-409]mm/yyyy"]];
        // E58
        const rangeDateDelivered = ws.getRange("E58");
        rangeDateDelivered.values = [[dateAuthorizationFormat]];
        rangeDateDelivered.numberFormat = [["dd/[$-409]mm/yyyy"]];
        // write medical examination
        const dateExaminationMedicalFormat = moment(this.medicalExamination).toOADate();
        const rangeDateExaminationMedical = ws.getRange("I8");
        rangeDateExaminationMedical.values = [[dateExaminationMedicalFormat]];
        rangeDateExaminationMedical.numberFormat = [["dd/[$-409]mm/yyyy"]];
        // write expiration date
        const dateExpirationFormat = moment(minExamDate).toOADate();
        const rangeDateExpiration = ws.getRange("S8");
        rangeDateExpiration.values = [[dateExpirationFormat]];
        rangeDateExpiration.numberFormat = [["dd/[$-409]mm/yyyy"]];
        // write lastname + firstName
        const rangeFullName = ws.getRange("D27");
        rangeFullName.values = [[this.userDataInfos.lastName + " " + this.userDataInfos.firstName]];
        // write responsible
        const rangeResponsible = ws.getRange("F28");
        rangeResponsible.values = [[this.responsible]];
        return context.sync();
      });
    }
  }
};
</script>