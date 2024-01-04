import moment from 'moment';

export const onFormatDate = (date?: Date): string => {
    return !date ? '' : moment(date).format('MM/DD/YYYY');
};
// export const onFormatDate = (date?: Date): string => {
//     return !date ? '' : date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear());
// };

export const formatSectionName = (text) => {
    if (text && text != "") {
        const sectionNameArr = text.trim().split(" ");
        if (sectionNameArr.length > 3) {
            return `${sectionNameArr[0]} ${sectionNameArr[1]} ${sectionNameArr[2]}`
        }
        else {
            return text;
        }
    }
    else
        return "";
}
export const mergeADSPUsers = (allADUsers, allSPUsers) => {
    let returnArr = allADUsers.map((adUser) => {
        let user = { ...adUser };
        const matchedSPUser = allSPUsers.filter((spUser) => spUser.Email == adUser.mail);
        if (matchedSPUser.length > 0) {
            user["SPUSerId"] = matchedSPUser[0].Id;
        }
        return user;
    })
    return returnArr;

}