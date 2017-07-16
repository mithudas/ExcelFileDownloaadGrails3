package exceldownload_grails3

import com.mithu.excelDownload.User

class BootStrap {

    def init = { servletContext ->
        createUser()
    }

    private def createUser() {
        println "creating users>>"
        
        new User(firstName: 'Ribhu', lastName: "Das", salary: 100000).save()
        new User(firstName: 'Imon', lastName: "Chowdhary", salary: 150000).save()
    }
    def destroy = {
    }
}
