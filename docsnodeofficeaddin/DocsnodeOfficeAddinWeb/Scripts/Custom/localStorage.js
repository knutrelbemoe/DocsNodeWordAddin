function setLocalForageItem(key, value) {
    var deferred = $.Deferred();
    try {
        localforage.setItem(key, value).then(function (value) {
            // Do other things once the value has been saved.
            deferred.resolve(value);
        }).catch(function (err) {
            // This code runs if there were any errors
            console.log(err);
        });
    } catch (err) {
        console.log("error setting item to localstorage : " + err);
    }
    return deferred.promise();
}

function getLocalForageItem(key) {

    var deferred = $.Deferred();
    try {
        localforage.getItem(key).then(function (value) {
            // This code runs once the value has been loaded
            // from the offline store.
            deferred.resolve(value);

        }).catch(function (err) {
            // This code runs if there were any errors
            console.log(err);
        });
    } catch (err) {
        console.log("error getting item from localstorage : " + err);
    }
    return deferred.promise();

}