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
      arrayCategoriesInfos: [
        {
          children: false,
          name: "workAtHeight",
          value: "",
          adressText: null,
          addressChecked: null,
          checked: false
        },
        {
          children: false,
          name: "AIPR",
          value: "",
          adressText: null,
          addressChecked: null,
          checked: false
        },
        {
          children: false,
          name: "shotCrete",
          value: "",
          adressText: null,
          addressChecked: null,
          checked: false
        },
        {
          children: false,
          name: "scaffoldingReception",
          value: "",
          adressText: null,
          addressChecked: null,
          checked: false
        },
        {
          children: false,
          name: "R482CatA",
          value: "",
          adressText: "C28",
          addressChecked: "B28",
          checked: false
        },
        {
          children: false,
          name: "R482CatB1",
          value: "",
          adressText: "C29",
          addressChecked: "B29",
          checked: false
        },
        {
          children: false,
          name: "R482CatB2",
          value: "",
          adressText: "C30",
          addressChecked: "B30",
          checked: false
        },
        {
          children: false,
          name: "R482CatB3",
          value: "",
          adressText: null,
          addressChecked: null,
          checked: false
        },
        {
          children: false,
          name: "R482CatC2",
          value: "",
          adressText: "C32",
          addressChecked: "B32",
          checked: false
        },
        {
          children: false,
          name: "R482CatC1",
          value: "",
          adressText: "C31",
          addressChecked: "B31",
          checked: false
        },
        {
          children: false,
          name: "R3725",
          value: "",
          adressText: null,
          addressChecked: null,
          checked: false
        },
        {
          children: false,
          name: "R482CatC3",
          value: "",
          adressText: "C33",
          addressChecked: "B33",
          checked: false
        },
        {
          children: false,
          name: "R482CatD",
          value: "",
          adressText: "C34",
          addressChecked: "B34",
          checked: false
        },
        {
          children: false,
          name: "R482CatE",
          value: "",
          adressText: "C35",
          addressChecked: "B35",
          checked: false
        },
        {
          children: false,
          name: "R482CatF",
          value: "",
          adressText: "C36",
          addressChecked: "B36",
          checked: false,
        },
        {
          children: false,
          name: "OptionWinch",
          value: "",
          adressText: "D37",
          addressChecked: "C37",
          checked: false
        },
        {
          children: false,
          name: "R482CatG",
          value: "",
          adressText: "C38",
          addressChecked: "B38",
          checked: false
        },
        {
          children: false,
          name: "OptionTMS",
          value: "",
          adressText: null,
          addressChecked: null,
          checked: false
        },
        {
          children: true,
          name: "R486",
          adressText: "C39:H39",
          addressChecked: "B39",
          nbRow: 3,
          checked: false,
          categories: [
            {
              name: "R486CatA1A",
              value: "",
              adressText: "F40",
              addressChecked: "E40",
              checked: false
            },
            {
              name: "R486CatB1B",
              value: "",
              adressText: "F41",
              addressChecked: "E41",
              checked: false
            },
            {
              name: "R486CatA3A",
              value: "",
              adressText: "H40",
              addressChecked: "G40",
              checked: false
            },
            {
              name: "R486CatB3B",
              value: "",
              adressText: "H41",
              addressChecked: "G41",
              checked: false
            }
          ]
        },
        {
          children: true,
          name: "R483",
          adressText: "C42:K44",
          addressChecked: "",
          nbRow: 3,
          checked: false,
          categories: [
            {
              name: "R483CatA1A",
              value: "",
              adressText: "I43",
              addressChecked: "H43",
              checked: false
            },
            {
              name: "R483CatB1B",
              value: "",
              adressText: "I44",
              addressChecked: "H44",
              checked: false
            },
            {
              name: "R483CatA2A",
              value: "",
              adressText: "K43",
              addressChecked: "J43",
              checked: false
            },
            {
              name: "R483CatB2B",
              value: "",
              adressText: "K44",
              addressChecked: "J44",
              checked: false
            }
          ]
        },
        {
          children: false,
          name: "R484Cat1",
          value: "",
          adressText: "M35",
          addressChecked: "L35",
          checked: false
        },
        {
          children: false,
          name: "R484Cat2",
          value: "",
          adressText: "M36",
          addressChecked: "L36",
          checked: false
        },
        {
          children: true,
          name: "R485",
          adressText: "M37:N40",
          addressChecked: "L37",
          nbRow: 3,
          checked: false,
          categories: [
            {
              name: "R485Cat1",
              value: "",
              adressText: "N38",
              addressChecked: "M38",
              checked: false
            },
            {
              name: "R485Cat2",
              value: "",
              adressText: "N39",
              addressChecked: "M39",
              checked: false
            }
          ]
        },
        {
          children: true,
          name: "R487",
          adressText: "M40:N43",
          addressChecked: "L40",
          nbRow: 4,
          checked: false,
          categories: [
            {
              name: "R487Cat1",
              value: "",
              adressText: "N41",
              addressChecked: "M41",
              checked: false
            },
            {
              name: "R487Cat2",
              value: "",
              adressText: "N42",
              addressChecked: "M42",
              checked: false
            },
            {
              name: "R487Cat3",
              value: "",
              adressText: "N43",
              addressChecked: "M43",
              checked: false
            }
          ]
        },
        {
          children: false,
          name: "R489Cat1A",
          value: "",
          adressText: "M28",
          addressChecked: "L28",
          checked: false
        },
        {
          children: false,
          name: "R489Cat1B",
          value: "",
          adressText: "M29",
          addressChecked: "L29",
          checked: false
        },
        {
          children: false,
          name: "R489Cat3",
          value: "",
          adressText: "M30",
          addressChecked: "L30",
          checked: false
        },
        {
          children: false,
          name: "R489Cat4",
          value: "",
          adressText: "M31",
          addressChecked: "L31",
          checked: false
        },
        {
          children: false,
          name: "R489Cat5",
          value: "",
          adressText: "M32",
          addressChecked: "L32",
          checked: false
        },
        {
          children: false,
          name: "R489Cat6",
          value: "",
          adressText: "M33",
          addressChecked: "L33",
          checked: false
        },
        {
          children: false,
          name: "R489Cat7",
          value: "",
          adressText: "M34",
          addressChecked: "L34",
          checked: false
        },
        {
          children: false,
          name: "R490",
          value: "",
          adressText: "M44",
          addressChecked: "L44",
          checked: false
        }
      ],
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
        // for init ibject, we create a loop on the array key
        for (let i = 0; i < this.keyDataDates.length; i++) {
          // we search the key corresponding in this object
          this.arrayCategoriesInfos.forEach(item => {
            // if it has children
            if(item.children) {
              // we search the child for define the value
              item.categories.forEach(elt => {
                if(elt.name === this.keyDataDates[i]) {
                  elt.value = rangeValues[0][i];
                }
              })
            } // otherwise, we search in this item for define the value
            else {
              if(item.name === this.keyDataDates[i]) {
                item.value = rangeValues[0][i];
              }
            }
          })
        }
        this.transformDateExcel("dates");
      } else {
        console.log('error type in initObject');
      }
    },
    // we transform the date excel into moment date
    transformDateExcel(type) {
      if(type === "dates") {
        this.arrayCategoriesInfos.forEach(item => {
            // if it has children
            if(item.children) {
              // we search the child for transform the value in moment
              item.categories.forEach(elt => {
                if(elt.value !== undefined && elt.value !== null && elt.value !== "") {
                  elt.value = moment.fromOADate(elt.value);
                }
              })
            } 
            // otherwise, we search in this item for transform the value in moment
            else {
              if(item.value !== undefined && item.value !== null && item.value !== "") {
                item.value = moment.fromOADate(item.value);
              }
            }
          });

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
        this.arrayCategoriesInfos.forEach(item => {
            // if it has children
            if(item.children) {
              // we search the child for checked the expiration
              item.categories.forEach(elt => {
                if(elt.value !== undefined && elt.value !== null && elt.value !== "") {
                  elt.checked = this.diffDate(elt.value);
                }
              })
            } 
            // otherwise, we search in this item for checked the expiration
            else {
              if(item.value !== undefined && item.value !== null && item.value !== "") {
                item.checked = this.diffDate(item.value);
              }
            }
          });
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
      this.writeUserDataDates();

      console.log(this.userDataInfos);
      console.log(this.arrayCategoriesInfos);
    },
    calculateMinDate() {
      const arrayValidDate = [];
      this.arrayCategoriesInfos.forEach(item => {
        // if it has children
        if(item.children) {
          // we search the checked children and push the date in array
          item.categories.forEach(elt => {
            if(elt.checked) {
              arrayValidDate.push(elt.value);
            }
          })
        } 
        // otherwise, we search the checked items and push the date in array
        else {
          if(item.checked) {
            arrayValidDate.push(item.value);
          }
        }
      });

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
    writeUserDataDates() {
      this.arrayCategoriesInfos.forEach((item) => {
        // if it has children
        if(item.children) {
          // we search the adressChecked children and call function for write data
          item.categories.forEach((elt) => {
            if(elt.addressChecked !== null && (item.name === "R483" || item.name === "R485" || item.name === "R486" || item.name === "R487")) {
              this.prepareUserDataDatesComplex(item);
              this.writeUserDataDatesLoop(elt.addressChecked, elt.checked);
            }
          })
        } 
        // otherwise, we search the adressChecked items and call function for write data
        else {
          if(item.addressChecked !== null) {
            this.writeUserDataDatesLoop(item.addressChecked, item.checked);
          }
        }
      });
    },
    prepareUserDataDatesComplex(item) {

        if(item.name === "R486") {
          // we define the arrays with different values for filter
          const categoryA = [ "R486CatA1A", "R486CatA3A" ];
          const categoryB = [ "R486CatB1B", "R486CatB3B" ];
          const category = categoryA.concat(categoryB);
          // we define if at least one is checked
          const isCategoryChecked = category.some(elt => item.categories[elt])
          const isCategoryAChecked = categoryA.some(elt => item.categories[elt])
          const isCategoryBChecked = categoryB.some(elt => item.categories[elt])
          // we call write function with address and value
          this.writeUserDataDatesLoop("B39", isCategoryChecked)
          this.writeUserDataDatesLoop("C40", isCategoryAChecked)
          this.writeUserDataDatesLoop("C41", isCategoryBChecked)
        } else if(item.name === "R483") {
          // we define the arrays with different values for filter
          const categoryA = [ "R483CatA1A", "R483CatA2A" ];
          const categoryB = [ "R483CatB1B", "R483CatB2B" ];
          const category = categoryA.concat(categoryB);
          // we define if at least one is checked
          const isCategoryChecked = category.some(elt => item.categories[elt])
          const isCategoryAChecked = categoryA.some(elt => item.categories[elt])
          const isCategoryBChecked = categoryB.some(elt => item.categories[elt])
          // we call write function with address and value
          this.writeUserDataDatesLoop("B42", isCategoryChecked)
          this.writeUserDataDatesLoop("C43", isCategoryAChecked)
          this.writeUserDataDatesLoop("C44", isCategoryBChecked)
        } else if(item.name === "R485") {
          // we define the arrays with different values for filter
          const category = [ "R485Cat1", "R485Cat2" ];
          // we define if at least one is checked
          const isCategoryChecked = category.some(elt => item.categories[elt])
          // we call write function with address and value
          this.writeUserDataDatesLoop("L37", isCategoryChecked)
        } else if(item.name === "R487") {
          // we define the arrays with different values for filter
          const category = [ "R487Cat1", "R487Cat2", "R487Cat3" ];
          // we define if at least one is checked
          const isCategoryChecked = category.some(elt => item.categories[elt])
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
        const range = ws.getRange(address);
        range.values = [[this.transformBooleanInNumber(value)]];

        return context.sync();
      })
    },
    transformBooleanInNumber(value) {
      return value === true ? 1 : 0;
    }
  }
};
</script>