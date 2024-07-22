function checkAllEventStatuses() {
    var calendarId = 'primary';
    var now = new Date();
    var twoWeeksFromNow = new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000);
    
    var events = Calendar.Events.list(calendarId, {
      timeMin: now.toISOString(),
      timeMax: twoWeeksFromNow.toISOString(),
      singleEvents: true,
      orderBy: 'startTime'
    });
    
    console.log("Found " + events.items.length + " total events");
    
    var myEmail = Session.getActiveUser().getEmail();
    
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      console.log("Event: " + event.summary);
      console.log("  My Status: " + (event.attendees ? event.attendees.find(a => a.self).responseStatus : "N/A"));
      console.log("  Organizer: " + (event.organizer ? event.organizer.email : "Unknown"));
      console.log("  Is Creator: " + (event.creator && event.creator.self));
      
      if (event.attendees) {
        console.log("  Attendees: " + event.attendees.map(a => a.email + " (" + a.responseStatus + ")").join(", "));
      }
      
      console.log("  Start: " + event.start.dateTime);
      console.log("  End: " + event.end.dateTime);
      
      var isInvitation = (event.organizer && event.organizer.email !== myEmail) || 
                         (event.attendees && event.attendees.find(a => a.self && a.responseStatus === "needsAction"));
      
      console.log("  Is Invitation: " + isInvitation);
      
      console.log("------------------");
    }
  }
  
  function autoHandleInvitations() {
    console.log("Script started running at " + new Date());
    
    var calendarId = 'primary';
    var now = new Date();
    var twoWeeksFromNow = new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000);
    
    var events = Calendar.Events.list(calendarId, {
      timeMin: now.toISOString(),
      timeMax: twoWeeksFromNow.toISOString(),
      singleEvents: true,
      orderBy: 'startTime'
    });
    
    console.log("Found " + events.items.length + " total events");
    
    var invitedEvents = 0;
    var myEmail = Session.getActiveUser().getEmail();
    
    for (var i = 0; i < events.items.length; i++) {
      var event = events.items[i];
      console.log("Checking event: " + event.summary);
      
      var isInvitation = (event.organizer && event.organizer.email !== myEmail) || 
                         (event.attendees && event.attendees.find(a => a.self && a.responseStatus === "needsAction"));
      
      if (isInvitation) {
        invitedEvents++;
        console.log("Found invited event: " + event.summary);
        
        var conflictingEvents = Calendar.Events.list(calendarId, {
          timeMin: event.start.dateTime,
          timeMax: event.end.dateTime,
          singleEvents: true
        });
        
        console.log("Found " + conflictingEvents.items.length + " potentially conflicting events");
        
        if (conflictingEvents.items.length > 1) {
          // There's a conflict
          var existingEvents = conflictingEvents.items.filter(e => e.id !== event.id);
          
          if (existingEvents.length > 0) {
            // The current event is the new invitation, so we'll decline it
            var updatedEvent = {
              attendees: event.attendees.map(attendee => {
                if (attendee.email === myEmail) {
                  attendee.responseStatus = "declined";
                }
                return attendee;
              })
            };
            
            Calendar.Events.patch(updatedEvent, calendarId, event.id);
            console.log("Declined new invitation: " + event.summary);
            
            var organizer = event.organizer ? event.organizer.email : (event.creator ? event.creator.email : null);
            
            if (organizer) {
              var conflictDetails = existingEvents.map(function(e) {
                var organizer = e.organizer ? e.organizer.email : (e.creator ? e.creator.email : "Unknown");
                var startDate = new Date(e.start.dateTime);
                var endDate = new Date(e.end.dateTime);
                return organizer + " has booked for " + 
                       formatDateTime(startDate) + " to " + 
                       formatTime(endDate);
              }).join("\n\n");
              
              GmailApp.sendEmail(organizer, 
                                 "Unable to attend: " + event.summary, 
                                 "I'm sorry, but I have conflicting events during this time slot. " +
                                 "Here are the details of the conflicting bookings:\n\n" + conflictDetails);
            }
          } else {
            // This shouldn't happen, but just in case, accept the invitation
            Calendar.Events.patch({
              attendees: [{ email: myEmail, responseStatus: "accepted" }]
            }, calendarId, event.id);
            console.log("Accepted invitation: " + event.summary);
          }
        } else {
          // No conflict, accept the invitation
          Calendar.Events.patch({
            attendees: [{ email: myEmail, responseStatus: "accepted" }]
          }, calendarId, event.id);
          console.log("Accepted invitation: " + event.summary);
        }
      }
    }
    
    console.log("Total invited events found: " + invitedEvents);
  }
  
  function formatDateTime(date) {
    return date.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }) + 
           ' at ' + formatTime(date);
  }
  
  function formatTime(date) {
    return date.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true });
  }