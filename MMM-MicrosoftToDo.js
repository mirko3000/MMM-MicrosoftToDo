Module.register("MMM-MicrosoftToDo",{

  	html: {
      empty: '<li style=\"list-style-position:inside; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;\">" + this.translate("NO_ENTRIES") + "</li>',
      table: '<thead>{0}</thead><tbody>{1}</tbody>',
      row: '<tr>{0}{1}</tr>',
      column: '<td width="6%">{0}</td><td width="2%">{1}</td><td class="title bright" align="left" width="40%">{2}</td><td width="2%">{3}</td>',
      star: '<i class="fa fa-star" aria-hidden="true"></i>',
      assignee: '<div style="display: inline-flex; align-items: center; justify-content: center; background-color: #aaa; color: #666; min-width: 1em; border-radius: 50%; vertical-align: middle; padding: 2px; text-transform: uppercase;">{0}</div>',
    },


    getDom: function () {

      var self = this;
      var wrapper = document.createElement("table");
      wrapper.className = "normal small light";

      var todos = this.list;

      var rows = []
      var rowCount = 0;
      var row;
      var previousCol;

      if (todos.length == 0) {
        wrapper.innerHTML = this.html.empty;
      }
      else {
        todos.forEach(function (todo, i) {

          if (i%2!=0) {
            var newCol = self.html.column.format(
              //self.config.showDeadline && todo.due_date ? todo.due_date : '',
              i+1,
              todo.importance == "high" ? self.html.star : '',
              todo.subject,
              self.config.showAssignee && todo.assignee_id && self.users ? self.html.assignee.format(self.users[todo.assignee_id]) : ''
            )
            rows[rowCount] = self.html.row.format(previousCol, newCol);
            previousCol = '';
            // New row
            rowCount++;
          }
          else {
            previousCol = self.html.column.format(
              //self.config.showDeadline && todo.due_date ? todo.due_date : '',
              i+1,
              todo.importance == "high" ? self.html.star : '',
              todo.subject,
              self.config.showAssignee && todo.assignee_id && self.users ? self.html.assignee.format(self.users[todo.assignee_id]) : ''
            )
          }

          // Create fade effect
          if (self.config.fade && self.config.fadePoint < 1) {
            if (self.config.fadePoint < 0) {
              self.config.fadePoint = 0;
            }
            var startingPoint = todos.length * self.config.fadePoint;
            if (i >= startingPoint) {
              wrapper.style.opacity = 1 - (1 / todos.length - startingPoint * (i - startingPoint));
            }
          }
        });

        if (previousCol != '') {
          rowCount++;
          var dummyCol = this.html.column.format('', '', '', '');
          rows[rowCount] = self.html.row.format(previousCol, dummyCol);
        }

        wrapper.innerHTML = this.html.table.format(
          this.html.row.format('', '', '', ''),
          rows.join('')
        )
      }

      return wrapper;
    },

    getTranslations: function() {
      return {
        en: "translations/en.js",
        de: "translations/de.js"
      }
    },

    socketNotificationReceived: function (notification, payload) {

      if (notification === ("FETCH_INFO_ERROR_" + this.config.id)) {

        console.error('An error occurred while retrieving the todo list from Microsoft To Do. Please check the logs.');
        console.error(payload.error);
        console.error(payload.errorDescription);
        this.list = [ { "subject": "Error occurred: " + payload.error + ". Check logs."} ];

        this.updateDom();

      }

      if (notification === ("DATA_FETCHED_" + this.config.id)) {

        this.list = payload;
        console.log(this.name + ' received list of ' + this.list.length + ' items.');

        // check if module should be hidden according to list size and the module's configuration
        if (this.config.hideIfEmpty) {
          if(this.list.length > 0) {
            if(this.hidden){
              this.show()
            }
          } else {
            if(!this.hidden) {
              console.log(this.name + ' hiding module according to \'hideIfEmpty\' configuration, since there are no tasks present in the list.');
              this.hide()
            }
          }
        }
      }

      this.updateDom()
    }
  },

  start: function () {
    // copy module object to be accessible in callbacks
    var self = this

    // start with empty list that shows loading indicator
    self.list = [{ subject: this.translate('LOADING_ENTRIES') }]

    // decide if a module should be shown if todo list is empty
    if (self.config.hideIfEmpty === undefined) {
      self.config.hideIfEmpty = false
    }

    // decide if a checkbox icon should be shown in front of each todo list item
    if (self.config.showCheckbox === undefined) {
      self.config.showCheckbox = true
    }

    // set default max module width
    if (self.config.maxWidth === undefined) {
      self.config.maxWidth = '450px'
    }

    // set default limit for number of tasks to be shown
    if (self.config.itemLimit === undefined) {
      self.config.itemLimit = '200'
    }

    // in case there are multiple instances of this module, ensure the responses from node_helper are mapped to the correct module
    self.config.id = this.identifier

    // update tasks every 60s
    var refreshFunction = function () {
      self.sendSocketNotification('FETCH_DATA', self.config)
    }
    refreshFunction()
    setInterval(refreshFunction, 60000)
  }

})
