from updater import Updater

if __name__ == '__main__':
    
    number_of_vendors = int(input('How many vendors do you need to update?\n'))
    updater = Updater()
    updater.prep_update()
    for i in range(number_of_vendors):
        updater.process_update()
    print('All vendors updated')
