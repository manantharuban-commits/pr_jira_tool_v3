-- ====================================================================
-- PR #42: feat: inventory & order management SQL overhaul
-- Generated : 2026-05-19 20:46:44
-- Files     : 3
-- ====================================================================

-- --------------------------------------------------------------------
-- Step 1: orders
-- File  : db/procedures/sp_update_order_status.sql
-- Status: MODIFIED  (+18 / -8)
-- --------------------------------------------------------------------

CREATE OR REPLACE PROCEDURE sp_update_order_status (
    p_order_id   IN  orders.order_id%TYPE,
    p_new_status IN  VARCHAR2,
    p_updated_by IN  VARCHAR2 DEFAULT USER,
    p_notes      IN  VARCHAR2 DEFAULT NULL,
    p_result     OUT VARCHAR2
) AS
    v_current_status  VARCHAR2(30);
    v_allowed_next    VARCHAR2(200);
    v_audit_id        audit_log.audit_id%TYPE;
BEGIN
    SELECT status
    INTO   v_current_status
    FROM   orders
    WHERE  order_id = p_order_id;

    IF p_new_status NOT IN ('PENDING','PROCESSING','PICKED','SHIPPED',
                            'OUT_FOR_DELIVERY','DELIVERED','CANCELLED','RETURNED') THEN
        p_result := 'ERROR: invalid status';
        RETURN;
    END IF;

    UPDATE orders
    SET    status       = p_new_status,
           updated_by   = p_updated_by,
           updated_date = SYSDATE
    WHERE  order_id = p_order_id;

    INSERT INTO order_status_history
           (order_id, old_status, new_status, changed_by, changed_date, notes)
    VALUES (p_order_id, v_current_status, p_new_status,
            p_updated_by, SYSDATE, p_notes);

    COMMIT;
    p_result := 'OK';
EXCEPTION
    WHEN OTHERS THEN
        ROLLBACK;
        p_result := 'ERROR: ' || SQLERRM;
END sp_update_order_status;


-- --------------------------------------------------------------------
-- Step 2: PACKAGE pkg_inventory_mgmt
-- File  : db/packages/pkg_inventory_mgmt.sql
-- Status: ADDED  (+72 / -0)
-- --------------------------------------------------------------------

CREATE OR REPLACE PACKAGE pkg_inventory_mgmt AS

    -- Reserve stock for an order line
    PROCEDURE reserve_stock (
        p_product_id  IN  NUMBER,
        p_qty         IN  NUMBER,
        p_order_id    IN  NUMBER,
        p_result      OUT VARCHAR2
    );

    -- Release reserved stock (cancellation / rejection)
    PROCEDURE release_stock (
        p_product_id  IN  NUMBER,
        p_qty         IN  NUMBER,
        p_order_id    IN  NUMBER,
        p_result      OUT VARCHAR2
    );

    -- Confirm stock deduction on dispatch
    PROCEDURE confirm_dispatch (
        p_order_id    IN  NUMBER,
        p_result      OUT VARCHAR2
    );

    -- Returns current available qty (on_hand - reserved)
    FUNCTION get_available_qty (
        p_product_id  IN  NUMBER
    ) RETURN NUMBER;

END pkg_inventory_mgmt;
/

CREATE OR REPLACE PACKAGE BODY pkg_inventory_mgmt AS

    PROCEDURE reserve_stock (p_product_id IN NUMBER, p_qty IN NUMBER,
                             p_order_id IN NUMBER, p_result OUT VARCHAR2) AS
        v_avail NUMBER;
    BEGIN
        SELECT on_hand_qty - reserved_qty
        INTO   v_avail
        FROM   inventory
        WHERE  product_id = p_product_id
        FOR UPDATE;

        IF v_avail < p_qty THEN
            p_result := 'ERROR: insufficient stock (' || v_avail || ' available)';
            RETURN;
        END IF;

        UPDATE inventory
        SET    reserved_qty = reserved_qty + p_qty
        WHERE  product_id = p_product_id;

        INSERT INTO inventory_reservations (product_id, order_id, qty, reserved_date)
        VALUES (p_product_id, p_order_id, p_qty, SYSDATE);

        COMMIT;
        p_result := 'OK';
    EXCEPTION WHEN OTHERS THEN ROLLBACK; p_result := 'ERROR: ' || SQLERRM;
    END reserve_stock;

    PROCEDURE release_stock (p_product_id IN NUMBER, p_qty IN NUMBER,
                             p_order_id IN NUMBER, p_result OUT VARCHAR2) AS
    BEGIN
        UPDATE inventory
        SET    reserved_qty = GREATEST(0, reserved_qty - p_qty)
        WHERE  product_id = p_product_id;

        DELETE FROM inventory_reservations
        WHERE  product_id = p_product_id AND order_id = p_order_id;

        COMMIT; p_result := 'OK';
    EXCEPTION WHEN OTHERS THEN ROLLBACK; p_result := 'ERROR: ' || SQLERRM;
    END release_stock;

    PROCEDURE confirm_dispatch (p_order_id IN NUMBER, p_result OUT VARCHAR2) AS
    BEGIN
        UPDATE inventory i
        SET    on_hand_qty  = on_hand_qty  - ir.qty,
               reserved_qty = reserved_qty - ir.qty
        FROM   inventory_reservations ir
        WHERE  ir.product_id = i.product_id AND ir.order_id = p_order_id;

        DELETE FROM inventory_reservations WHERE order_id = p_order_id;
        COMMIT; p_result := 'OK';
    EXCEPTION WHEN OTHERS THEN ROLLBACK; p_result := 'ERROR: ' || SQLERRM;
    END confirm_dispatch;

    FUNCTION get_available_qty (p_product_id IN NUMBER) RETURN NUMBER AS
        v_qty NUMBER := 0;
    BEGIN
        SELECT on_hand_qty - NVL(reserved_qty, 0)
        INTO   v_qty FROM inventory WHERE product_id = p_product_id;
        RETURN v_qty;
    EXCEPTION WHEN NO_DATA_FOUND THEN RETURN 0;
    END get_available_qty;

END pkg_inventory_mgmt;


-- --------------------------------------------------------------------
-- Step 3: v_customer_summary
-- File  : db/views/v_customer_summary.sql
-- Status: MODIFIED  (+3 / -2)
-- --------------------------------------------------------------------

CREATE OR REPLACE VIEW v_customer_summary AS
SELECT
    c.customer_id,
    c.full_name,
    c.email,
    c.created_date,
    COUNT(o.order_id)       AS total_orders,
    NVL(lp.tier, 'BRONZE')  AS loyalty_tier
FROM customers c
LEFT JOIN orders o         ON o.customer_id  = c.customer_id
LEFT JOIN loyalty_points lp ON lp.customer_id = c.customer_id
WHERE c.status = 'ACTIVE'
GROUP BY
    c.customer_id, c.full_name, c.email,
    c.created_date, lp.tier;

