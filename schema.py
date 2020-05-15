import yaml
import pathlib
import re
from pprint import pprint

class Schema:
    site = list() 
    playbooks = list()
    groups = list()
    hosts = list() 
    globalVars = dict()

    def __loadAnsibleVars(self, targetDir, varsFile):
        targetPath = targetDir / varsFile
        if targetPath.is_dir():
            result = sorted(targetPath.glob('**/*.yml'))
        elif (targetDir / (varsFile + '.yml')).is_file():
            result = [(targetDir / (varsFile + '.yml'))] 
        else:
            result = []

        return result

    def load(self, 
                 playbook_path, 
                 environment='production'):
        rootPath = pathlib.Path(playbook_path)
        groupVarsPath = rootPath / 'group_vars' 
        hostVarsPath = rootPath / 'host_vars' 

        # Load site.yml
        with (rootPath / 'site.yml').open('r') as f:
            site = yaml.safe_load(f)
            for item in site:
                if 'import_playbook' in item:
                    self.site.append(item['import_playbook'])

        # Load playbooks.
        for playbookPath in self.site:
            with (rootPath / playbookPath ).open('r') as f:
                playbook = yaml.safe_load(f)[0]
                playbook['path'] = playbookPath 
                self.playbooks.append(playbook)

        # Load Global vars
        globalVarsFiles = self.__loadAnsibleVars(
                            rootPath / 'group_vars',
                            'all') 
        for path in globalVarsFiles:
            with path.open('r') as f:
                globalVars = yaml.safe_load(f)
                self.globalVars.update(globalVars)

        # Load environment file.
        with (rootPath / environment).open('r') as f:
            groupPattern = re.compile('\[.*\]')
            hostPattern = re.compile('^[a-z,A-Z,0-9]')
            group=dict()
            for line in f.readlines():
                if groupPattern.match(line):
                    group = {
                        "name": re.sub('(\[|\])', '', line.strip()),
                        "hosts": list()
                    }
                    self.groups.append(group)
                elif hostPattern.match(line):
                    group['hosts'].append(line.strip())
        
        # Load group vars.
        for group in self.groups:
            groupName = group['name']
            group['vars'] = dict()
            groupVarsFiles = self.__loadAnsibleVars(
                                rootPath / 'group_vars',
                                groupName
                            )
            for path in groupVarsFiles:
                with path.open('r') as f:
                    groupVars = yaml.safe_load(f)
                    group['vars'].update(groupVars) 

        # Load Host
        tempList = list()
        for group in self.groups:
            tempList.extend(group['hosts'])
        for hostName in list(dict.fromkeys(tempList)):
            host = dict()
            host['name'] = hostName
            host['vars'] = dict()
            hostVarsFiles = self.__loadAnsibleVars(
                                rootPath / 'host_vars',
                                hostName
                            )
            for path in hostVarsFiles:
                with path.open('r') as f:
                    hostVars = yaml.safe_load(f)
                    host['vars'].update(hostVars)

            self.hosts.append(host)


    def getSite(self):
        return self.site

    def getPlaybooks(self):
        return self.playbooks

    def getGroups(self):
        return self.groups

    def getHosts(self):
        return self.hosts

    def getGlobalVars(self):
        return self.globalVars

if __name__ == "__main__":
    schema = Schema()
    schema.load("./ansible-playbook")
    pprint(schema.getSite())
    pprint(schema.getPlaybooks())
    pprint(schema.getGroups())
    pprint(schema.getHosts())
    pprint(schema.getGlobalVars())
            
