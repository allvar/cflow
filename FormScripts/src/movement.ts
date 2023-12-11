// --------------------------------------------------------------------------------------------
// --------------------------------------------------------------------------------------------
// onLoad - EVENT HANDLERS
// --------------------------------------------------------------------------------------------
export async function onLoad(executionContext: Xrm.Events.EventContext) {
  const formContext: Xrm.FormContext = executionContext.getFormContext();
  const formType = formContext.ui.getFormType();

  if (formType === XrmEnum.FormType.Create) {
    const luCustomer = formContext
      .getAttribute<Xrm.Attributes.LookupAttribute>("cflow_customer_id")
      ?.getValue();

    if (luCustomer) {
      const addressId = await setPrimaryAddress(formContext, luCustomer[0].id);
      if (addressId) setAddress(formContext, addressId);
    }
  }
}

// --------------------------------------------------------------------------------------------
// --------------------------------------------------------------------------------------------
// onChange - EVENT HANDLERS
// --------------------------------------------------------------------------------------------

export async function customer_id_onChange(
  executionContext: Xrm.Events.EventContext
) {
  const formContext: Xrm.FormContext = executionContext.getFormContext();
  const luCustomer = formContext
    .getAttribute<Xrm.Attributes.LookupAttribute>("cflow_customer_id")
    ?.getValue();

  if (luCustomer) {
    const addressId = await setPrimaryAddress(formContext, luCustomer[0].id);
    if (addressId) setAddress(formContext, addressId);
  }
}

/**
 * Field onChange event for "address_id", i.e. change of address lookup
 * - Sets three address fields to address of selected address
 */
export async function address_id_onChange(
  executionContext: Xrm.Events.EventContext
) {
  const formContext: Xrm.FormContext = executionContext.getFormContext();
  const luAddress = formContext
    .getAttribute<Xrm.Attributes.LookupAttribute>("cflow_address_id")
    ?.getValue();

  if (luAddress) {
    setAddress(formContext, luAddress[0].id);
  }
}

// --------------------------------------------------------------------------------------------
// --------------------------------------------------------------------------------------------
// SUPPORT FUNCTIONS
// --------------------------------------------------------------------------------------------

/**
    Support function to set address fields to address of address entity
    @param formContext
    @param addressId ID of address entity
  */
function setAddress(formContext: Xrm.FormContext, addressId: string) {
  Xrm.WebApi.retrieveRecord("cflow_address", addressId).then(
    (result) => {
      formContext
        .getAttribute("cflow_address_line_1")
        .setValue(result.cflow_address_line_1);
      formContext
        .getAttribute("cflow_address_line_2")
        .setValue(result.cflow_address_line_2);
      formContext
        .getAttribute("cflow_address_line_3")
        .setValue(result.cflow_address_line_3);
    },
    (error) => {
      console.log("Error at fetch address entity", error);
    }
  );
}

/**
    Support function to set address lookup to primary address from customer
    @param formContext
    @param customerId ID of customer entity
    @returns ID of primary address
  */
async function setPrimaryAddress(
  formContext: Xrm.FormContext,
  customerId: string
): Promise<string | undefined> {
  return new Promise((resolve, reject) => {
    Xrm.WebApi.retrieveRecord("cflow_customer", customerId).then(
      (result) => {
        formContext.getAttribute("cflow_address_id").setValue([
          {
            id: result._cflow_primary_address_id_value,
            name: result[
              "_cflow_primary_address_id_value@OData.Community.Display.V1.FormattedValue"
            ],
            entityType: "cflow_address",
          },
        ]);

        resolve(result._cflow_primary_address_id_value);
      },
      (error) => {
        console.log("Error at fetch customer entity", error);
        reject(undefined);
      }
    );
  });
}
