<template>
  <div id="app" class="bg-gray-200">
    <!-- Header -->
    <div class="p-12 bg-blue-400 flex justify-center items-center">
      <h1 class="text-xl text-gray-50">
        Autorisation de conduite
      </h1>
    </div>
    <!-- general informations --> 
    <div class="bg-gray-50 p-4">
      <div>
        <h2 class="text-md font-normal">
          Informations genérales:
        </h2>
        <!-- Responsible input -->
        <div>
          <label for="responsible" class="text-sm font-thin">Nom et Prénom du Délégant :</label>
          <input id="responsible" type="text" v-model="responsible" class="bg-transparent" required/>
        </div>
        <!-- limit date input -->
        <div>
          <label for="limitDate" class="text-sm font-thin">Date d'expiration (en jrs) :</label>
          <input id="limitDate" type="number" min="0" placeholder="0" v-model="limitDate" class="bg-transparent" required/>
        </div>
      </div>
      <!-- select user row input -->
      <div class="my-2">
        <h2 class="text-md font-normal">
          Conducteur/trice d'engins :
        </h2>
        <label for="row" class="text-sm font-thin">Entrer le numéro de ligne de la personne :</label>
        <input id="row" type="number" min="3" placeholder="3" v-model="rowInput" class="bg-transparent" @change="readData()" />
      </div>
      <!-- user selected and button for edit document -->
      <div v-if="userDataChecked" class="rounded-md border border-blue-500 flex flex-col items-center mt-4 p-4">
        <p>Utilisateur sélectionné :</p>
        <p v-if="userDataInfos.isPresent">
          {{ userDataInfos.lastName }} {{ userDataInfos.firstName }}
        </p>
        <button v-if="userDataInfos.isPresent" class="bg-blue-600 text-gray-200 py-2 px-3 rounded-md transition duration-500 hover:bg-blue-800 hover:text-gray-50 mt-4" @click="editDocument">
          Générer le document
        </button>
        <p v-else>
          {{ userDataInfos.lastName }} {{ userDataInfos.firstName }} ne fait plus parti des effectifs.
        </p>
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
      responsible: null,
      limitDate: 7,
      rowInput: 3,
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

        console.log(this.userDataInfos);
      } else if (type === "dates") {
        this.userDataDates = {};

        for (let i = 0; i < this.keyDataDates.length; i++) {
          this.userDataDates[this.keyDataDates[i]] = rangeValues[0][i];
        }

        console.log(this.userDataDates);
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
        console.log('transformDateExcel dates', this.userDataDates);
        this.isTheDateValid(type);

      } else if(type === "medical") {
        if(this.medicalExamination !== "") {
          this.medicalExamination = moment.fromOADate(this.medicalExamination);
          this.isTheDateValid(type);
        } else {
          this.medicalExamination = null;
        }
        console.log('transformDateExcel medical', this.medicalExamination);
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
        console.log(this.userDataDatesChecked);
      } else if(type === "medical") {
        this.medicalExaminationChecked = this.diffDate(this.medicalExamination);
        console.log('medicalChecked', this.medicalExaminationChecked);
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
      this.writeUserDataInfos();
    },
    writeUserDataInfos() {
      window.Excel.run((context) => {
        const ws = context.workbook.worksheets.getItem('TEMPLATE');
        const range = ws.getRange("I8");
        range.values = [[44]];
        return context.sync();
      });
    }
  }
};
</script>