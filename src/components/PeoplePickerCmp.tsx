import { PeoplePicker, PersonType } from "@microsoft/mgt-react";

function PeoplePickerCmp() {
  const userData = ["bikash.bagaria0@cfexperian.com"];

  function onChangePeople(e: any) {
    console.log(e.detail);
  }

  return (
    <>
      <div className="person-cls">
        <PeoplePicker
          defaultSelectedUserIds={userData}
          userIds={userData}
          type={PersonType.person}
          selectionMode="multiple"
          selectionChanged={onChangePeople}
        />
      </div>
    </>
  );
}

export default PeoplePickerCmp;
